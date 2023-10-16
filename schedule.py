"""
This implements the Webex Webinar scheduling process.
Can be launched as a standalone script or invoked by the control bot command.
"""

# common imports
import os
from dotenv import load_dotenv
import logging
import io
import json
from datetime import datetime, timedelta, timezone
from email.utils import getaddresses
import tempfile

# SDK imports
import webexteamssdk
import smartsheet

# my imports
from param_store import (
    getSmartsheetId,
    getWebexIntegrationToken
)
from exceptions import (
    ParameterStoreError,
    SmartsheetInitError,
    SmartsheetColumnMappingError,
    WebexIntegrationInitError,
    WebexBotInitError
)


def loadParameters():
    """
        First step in the scheduling process.

        Loads credentials and other parameters from env variables.
        Loads other parameters from Parameter Store.
        Checks if all required values are provided.
        Changes global variables.

        Args:
            None

        Returns:
            None
    """
    # Required user-set parameters from the env
    global SMARTSHEET_ACCESS_TOKEN
    global WEBEX_INTEGRATION_CLIENT_ID, WEBEX_INTEGRATION_CLIENT_SECRET
    global WEBEX_BOT_TOKEN, WEBEX_BOT_ROOM_ID
    # Optional user-set parameters from the env
    global SMARTSHEET_PARAMS, WEBEX_INTEGRATION_PARAMS

    load_dotenv(override=True)

    # load required parameters from env
    logger.info("Loading required parameters from env.")
    SMARTSHEET_ACCESS_TOKEN = os.getenv('SMARTSHEET_ACCESS_TOKEN')
    if not SMARTSHEET_ACCESS_TOKEN:
        logger.fatal("Smartsheet access token is missing. Provide with SMARTSHEET_ACCESS_TOKEN environment variable.")
        raise SystemExit()
    WEBEX_INTEGRATION_CLIENT_ID = os.getenv('WEBEX_INTEGRATION_CLIENT_ID')
    if not WEBEX_INTEGRATION_CLIENT_ID:
        logger.fatal("Webex Integration Client ID is missing. Provide with WEBEX_INTEGRATION_CLIENT_ID environment variable.")
        raise SystemExit()
    WEBEX_INTEGRATION_CLIENT_SECRET = os.getenv('WEBEX_INTEGRATION_CLIENT_SECRET')
    if not WEBEX_INTEGRATION_CLIENT_SECRET:
        logger.fatal("Webex Integration Client Secret is missing. Provide with WEBEX_INTEGRATION_CLIENT_SECRET environment variable.")
        raise SystemExit()
    WEBEX_BOT_TOKEN = os.getenv('WEBEX_BOT_TOKEN')
    if not WEBEX_BOT_TOKEN:
        logger.fatal("Webex Bot access token is missing. Provide with WEBEX_BOT_TOKEN environment variable.")
        raise SystemExit()
    WEBEX_BOT_ROOM_ID = os.getenv('WEBEX_BOT_ROOM_ID')
    if not WEBEX_BOT_ROOM_ID:
        logger.fatal("Webex Bot room ID is missing. It is required for logging and control. Provide with WEBEX_BOT_ROOM_ID environment variable.")
        raise SystemExit()
    logger.info("Required parameters are loaded from env.")

    # load optional parameters from env
    SMARTSHEET_PARAMS = {}
    SMARTSHEET_PARAMS['columns'] = {
        'create': "Create",
        'startdate': "Start Date",
        'starttime': "Start Time",
        'duration': "Duration",
        'title': "Title",
        'agenda': "Agenda",
        'cohosts': "Cohosts",
        'panelists': "Panelists",
        'webinarId': "Webinar ID",
        'attendeeUrl': "Attendee URL",
        'hostKey': "Host Key",
        'registrantCount': "Registrant Count"

    }
    SMARTSHEET_PARAMS['nicknames'] = {}
    if os.getenv("SMARTSHEET_PARAMS"):
        logger.info("Loading optional Smartsheet parameters from env.")
        try:
            columns = json.loads(os.getenv('SMARTSHEET_PARAMS'))['columns']
            for i in columns:
                SMARTSHEET_PARAMS['columns'][i] = columns[i]
            logger.info("Optional Smartsheet column parameters are loaded from env.")
        except Exception:
            logger.info("Could not load optional Smartsheet column parameters.")
        try:
            SMARTSHEET_PARAMS['nicknames'] = json.loads(os.getenv('SMARTSHEET_PARAMS'))['nicknames']
            logger.info("Optional Smartsheet nickname parameters are loaded from env.")
        except Exception:
            logger.info("Could not load optional Smartsheet nickname parameters.")

    else:
        logger.info("No optional Smartsheet parameters set in env.")

    WEBEX_INTEGRATION_PARAMS = {}
    if os.getenv('WEBEX_INTEGRATION_PARAMS'):
        logger.info("Loading optional Webex Integration parameters from env.")
        try:
            WEBEX_INTEGRATION_PARAMS = json.loads(os.getenv('WEBEX_INTEGRATION_PARAMS'))
            logger.info("Optional Webex Integration parameters are loaded from env.")
        except Exception as ex:
            logger.info("Could not load optional Webex Integration parameters. " + str(ex))
    else:
        logger.info("No optional Webex Integration parameters set in env.")


def initSmartsheet():
    """Initializes access to Smartsheet.

        Args:
            None

        Returns:
            A tuple of three items (ssApi, ssSheet, ssColumnMap)
                ssApi: Smartsheet API connection object
                ssSheets: Smartsheet sheet data model
                ssColumnMap: dict mapping column names to Smartsheet column IDs

        Raises:
            ParameterStoreErrorr: Could not read Smartsheet ID from parameter store.
            SmartsheetInitError: An error occurred accessing Smartsheet or Smartsheet ID is wrong.
            SmartsheetColumnMappingError: Smartsheet columns could not be properly mapped.
             Occurs when one or more of the required columns are missing from the provided smartsheet.
    """
    try:
        sheetId = getSmartsheetId()
    except Exception as ex:
        raise ParameterStoreError("Could not read Smartsheet ID from parameter store. " + str(ex))

    try:
        ssApi = smartsheet.Smartsheet()
        ssApi.errors_as_exceptions(True)
        ssColumnMap = {}

        ssSheet = ssApi.Sheets.get_sheet(sheetId, level=2, include=['objectValue'])
        # level=1 and include=['objectValue'] are required to receive more detailed responses from Smartsheet API, including MULTI_CONTACT
    except Exception as ex:
        raise SmartsheetInitError("Could not initialize Smartsheet. Check if Smartsheet ID is correct. " + str(ex))

    columns = {}
    # Build column map for later reference - translates column names to column id
    for column in ssSheet.columns:
        columns[column.title] = column.id
    for column in SMARTSHEET_PARAMS['columns']:
        if SMARTSHEET_PARAMS['columns'][column] in columns:
            ssColumnMap[column] = columns[SMARTSHEET_PARAMS['columns'][column]]
    # check if all required columns are present in the smartsheet
    requiredColumns = ['create', 'startdate', 'starttime', 'title', 'webinarId']
    columnsDiff = set(requiredColumns) - set(ssColumnMap.keys())
    if columnsDiff:
        raise SmartsheetColumnMappingError("Some required column(s) is(are) missing in your smartsheet: " + ", ".join(columnsDiff))

    return (ssApi, ssSheet, ssColumnMap)


def initWebexIntegration():
    """Initializes access to Webex integration.

        Args:
            None

        Returns:
            webexApi: webexteamssdk.WebexTeamsAPI object to control access to Webex Integration

        Raises:
            WebexIntegrationInitError
    """
    try:
        # get a fresh token
        webexToken = getWebexIntegrationToken(WEBEX_INTEGRATION_CLIENT_ID, WEBEX_INTEGRATION_CLIENT_SECRET)
        print(webexToken) #TODO dev
        # init the object
        webexApi = webexteamssdk.WebexTeamsAPI(webexToken)
        # check if API is functional by requesting `me`
        assert webexApi.people.me()
    except Exception as ex:
        raise WebexIntegrationInitError(ex)

    return webexApi


def initWebexBot():
    """Initializes access to Webex bot for logging and control.

        Args:
            None

        Returns:
            botApi: webexteamssdk.WebexTeamsAPI object to control access to Webex Integration
    """
    try:
        botApi = webexteamssdk.WebexTeamsAPI(WEBEX_BOT_TOKEN)
        # check if it's a bot
        assert botApi.people.me().type == "bot"
        # check if the bot can access the logging room
        assert botApi.rooms.get(WEBEX_BOT_ROOM_ID)
    except Exception:
        raise WebexBotInitError()

    return botApi


def getWebinarProperty(propertyName, ssRow=None):
    """Returns value for webinar property taken from a source, in order of piority:
        1. If propertyName-associated column exists in Smartsheet and cell is not empty,
            return its value
        2. If propertyName-associated filed exists in WEBEX_INTEGRATION_PARAMS and not empty,
            return its value
        3. Return None
    Provides special treatment to non-text fields (cohosts, panelists).

        Args:
            propertyName: Name of the Webinar property. Naming aligns with API specs
            ssRow: Smartsheet row

        Returns:
            propertyValue (str): for number/text fields,
            propertyValue (list): for object fields, i.e. MULTI_CONTACT
            or None
    """
    # if property value is specified in Smartsheet, return it
    try:
        cell = ssRow.get_column(ssColumnMap[propertyName])
        if cell.object_value.object_type == smartsheet.models.object_value.MULTI_CONTACT:
            propertyValue = dict((i.email.strip(), i.name.strip()) for i in cell.object_value.values)
            return propertyValue
        else:
            propertyValue = str(cell.value).strip()
            return propertyValue
    except Exception:
        pass
    # if property value is specified in env WEBEX_INTEGRATION_PARAMS, return it
    try:
        propertyValue = WEBEX_INTEGRATION_PARAMS[propertyName]
        return propertyValue
    except Exception:
        pass
    # otherwise, return None
    return None


def stringContactsToDict(contacts):
    """Returns dict of contacts from string of contacts

        Args: contacts(str): Comma-separated list of contacts, each contact represented as "name <email>".
            If name is not specified, 'Panelist' is used.
            If email is not specified, the function will try to match contact by nickname configured in env.

        Returns:
            _res: dict of contacts {email: name}

    """
    _res = {}
    for contact in getaddresses(str(contacts).split(",")):
        if "@" in contact[1]:
            # email specified
            _res[contact[1].strip().lower()] = contact[0].strip() or "Panelist"
        else:
            # email not specified
            try:
                _res[SMARTSHEET_PARAMS['nicknames'][contact[1].strip().lower()]['email']] = SMARTSHEET_PARAMS['nicknames'][contact[1].strip().lower()]['name']
            except Exception:
                pass

    return _res


if __name__ == "__main__":

    #
    #   Initialize logging
    #
    logger = logging.getLogger(__name__)
    # Logger usage:
    # logger.fatal("Message in case of a fatal error causing SystemExit")
    # logger.error("Message in case of an error, goes to the brief log")
    # logger.warning("Message to be output to the brief log (the Webex message itself)")
    # logger.info("Message to be output in the full log (text file attached to Webex message)")
    # logger.debug("Message to be output in the console only")

    # set log level to DEBUG
    logger.setLevel(logging.DEBUG)

    # log handler for brief log
    briefLogString = io.StringIO()
    briefLogHandler = logging.StreamHandler(briefLogString)
    briefLogHandler.setLevel(logging.WARNING)
    logger.addHandler(briefLogHandler)

    # log handler for full log
    fullLogString = io.StringIO()
    fullLogHandler = logging.StreamHandler(fullLogString)
    fullLogHandler.setLevel(logging.INFO)
    logger.addHandler(fullLogHandler)

    # log handler for console
    consoleLogHandler = logging.StreamHandler()
    consoleLogHandler.setLevel(logging.DEBUG)
    logger.addHandler(consoleLogHandler)

    logger.warning("Starting...")

    #
    # Load env variables and check if all env variables are provided
    #
    logger.info("Loading parameters and checking if all required parameters are provided.")
    loadParameters()
    logger.info("Required parameters are successfully loaded.")

    #
    # Initialize access to Smartsheet
    #
    logger.info("Initializing access to Smartsheet.")
    try:
        ssApi, ssSheet, ssColumnMap = initSmartsheet()
    except ParameterStoreError as ex:
        logger.fatal("Could not read Smartsheet Sheet ID from Parameter Store. Check local AWS configuration. Service reported: " + str(ex))
        raise SystemExit()
    except SmartsheetInitError as ex:
        logger.fatal("Smartsheet API connection error. " + str(ex))
        raise SystemExit()
    except SmartsheetColumnMappingError as ex:
        logger.fatal("Smartsheet column mapping error. " + str(ex))
        raise SystemExit()
    except Exception as ex:
        logger.fatal("Smartsheet initialization error. " + str(ex))
        raise SystemExit()
    logger.info("Successfully initialized access to Smartsheet.")

    #
    # Initialize access to Webex Integration
    #
    logger.info("Initializing access to Webex Integration.")
    try:
        webexApi = initWebexIntegration()
    except ParameterStoreError as ex:
        logger.fatal("Could not read Webex Integration tokens from Parameter Store. Check local AWS configuration. Service reported: " + str(ex))
        raise SystemExit()
    except WebexIntegrationInitError as ex:
        logger.fatal("Could not initialize Webex Integration. Service reported: " + str(ex))
        raise SystemExit()
    except Exception as ex:
        logger.fatal("Could not initialize Webex Integration. Service reported: " + str(ex))
        raise SystemExit()
    logger.info("Successfully initialized access to Webex Integration.")

    #
    # Initialize access to Webex bot for logging and control
    #
    logger.info("Initializing access to Webex bot.")
    try:
        botApi = initWebexBot()
    except Exception as ex:
        logger.fatal("Could not initialize Webex bot. Service reported: " + str(ex))
        raise SystemExit()
    logger.info("Successfully initialized access to Webex bot.")


    #
    # Set default time zone
    #
    os.environ['TZ'] = 'UTC'


    #
    # Loop over the smartsheet
    #
    for ssRow in ssSheet.rows:

        if str(ssRow.get_column(ssColumnMap['create']).value).lower() == "yes":
            event = {}

            logger.info("")    # insert empty line into log

            # gather all webinar properties
            event['title'] = getWebinarProperty('title', ssRow) or "Generic Webinar Title"
            try:
                event['agenda'] = getWebinarProperty('agenda', ssRow)
                event['scheduledType'] = getWebinarProperty('scheduledType', ssRow) or 'webinar'
                event['startdatetime'] = datetime.strptime(
                    ssRow.get_column(ssColumnMap['startdate']).value
                    + " " +
                    ssRow.get_column(ssColumnMap['starttime']).value,
                    "%Y-%m-%d %H:%M"
                )
                event['startdatetime'] = event['startdatetime'].replace(tzinfo=timezone.utc)    # set Zulu time (UTC timezone), Smartsheet always returns dates in UTC
                event['duration'] = getWebinarProperty('duration', ssRow) or 60    # by default, set duration to 1 hour
                event['duration'] = int(float(event['duration']))    # make sure it's integer
                event['enddatetime'] = event["startdatetime"] + timedelta(minutes=event['duration'])
                event['timezone'] = getWebinarProperty('timezone', ssRow) or "UTC"
                event['siteUrl'] = getWebinarProperty('siteUrl', ssRow)    # if not set, a Webex default will be used
                event['password'] = getWebinarProperty('password', ssRow)    # by default, randomly generated by Webex
                event['panelistPassword'] = getWebinarProperty('panelistPassword', ssRow)    # by default, randomly generated by Webex
                event['reminderTime'] = getWebinarProperty('reminderTime', ssRow) or 30    # by default, set reminder to go 30 minutes before the session
                # if it's too late to send reminder, skip it
                if datetime.utcnow().replace(tzinfo=timezone.utc) >= event['startdatetime'] - timedelta(minutes=event['reminderTime']):
                    event['reminderTime'] = 0
                event['registration'] = getWebinarProperty('registration', ssRow) or \
                    {
                        'autoAcceptRequest': True,
                        'requireFirstName': True,
                        'requireLastName': True,
                        'requireEmail': True
                    }    # registration is enabled by default
                event['enabledJoinBeforeHost'] = getWebinarProperty('enabledJoinBeforeHost', ssRow)   # let attendees join before host
                event['joinBeforeHostMinutes'] = getWebinarProperty('joinBeforeHostMinutes', ssRow)   # set webinar to start minutes before the scheduled start time

                # add invited cohosts
                event['cohosts'] = getWebinarProperty('cohosts', ssRow)
                if not isinstance(event['cohosts'], dict):
                    event['cohosts'] = stringContactsToDict(event['cohosts'])
                # add invited panelists
                event['panelists'] = getWebinarProperty('panelists', ssRow)
                if not isinstance(event['panelists'], dict):
                    event['panelists'] = stringContactsToDict(event['panelists'])
                # add panelists which are always invited
                alwaysInvitePanelists = getWebinarProperty('alwaysInvitePanelists')
                alwaysInvitePanelists = stringContactsToDict(alwaysInvitePanelists)
                event['panelists'].update(alwaysInvitePanelists)

                event['id'] = ssRow.get_column(ssColumnMap['webinarId']).value
                logger.info("Processing \"{}\"".format(event['title']))
            except Exception as ex:
                logger.error("Failed to process \"{}\". The webinar property is not valid: {}".format(event['title'], ex))
                continue

            # dev
            # import random
            # _rnd = "".join(random.choice([chr(c) for c in range(ord('A'), ord('Z')+1)]) for l in range(3))
            # event["title"] = _rnd+" "+event["title"]
            # /dev

            if not event['id']:
                # create event
                try:
                    w = webexApi.meetings.create(
                        title=event['title'],
                        agenda=event['agenda'],
                        scheduledType=event['scheduledType'],
                        start=str(event["startdatetime"]),
                        end=str(event["enddatetime"]),
                        timezone=event['timezone'],
                        siteUrl=event['siteUrl'],
                        password=event['password'],
                        panelistPassword=event['panelistPassword'],
                        reminderTime=event['reminderTime'],
                        registration=event['registration'],
                        enabledJoinBeforeHost=event['enabledJoinBeforeHost'],
                        joinBeforeHostMinutes=event['joinBeforeHostMinutes']
                    )
                    logger.warning("Created webinar {}".format(w.title))
                except Exception as ex:
                    logger.error("Failed to create webinar \"{}\". API returned error: {}".format(event['title'], ex))
                    try:
                        for err in ex.details['errors']:
                            logger.error("  " + err['description'])
                    except Exception:
                        pass
                    continue
                pass

                # update newly created webinar ID and info back into Smartsheet
                try:
                    newCells = []
                    if 'webinarId' in ssColumnMap:
                        newCells.append(ssApi.models.Cell({
                            'column_id': ssColumnMap['webinarId'],
                            'value': w.id
                        }))
                    else:
                        logger.error("No column in Smartsheet to save Webinar ID.")    # critical for app logic
                    if 'attendeeUrl' in ssColumnMap:
                        newCells.append(ssApi.models.Cell({
                            'column_id': ssColumnMap['attendeeUrl'],
                            'value': "Manually copy the Attendee URL from Webex" #w.webLink     TODO: implement Attendee URL once it becomes available in API
                        }))
                    else:
                        logger.info("No column in Smartsheet to save Attendee URL.")
                    # TODO to be added when API supports - Registration URL
                    if 'hostKey' in ssColumnMap:
                        newCells.append(ssApi.models.Cell({
                            'column_id': ssColumnMap['hostKey'],
                            'value': w.hostKey
                        }))
                    else:
                        logger.info("No column in Smartsheet to save Host Key.")

                    newRow = ssApi.models.Row()
                    newRow.id = ssRow.id
                    newRow.cells.extend(newCells)
                    ss = ssApi.Sheets.update_rows(ssSheet.id, [newRow])
                    logger.info("Updated webinar information into Smartsheet.")
                except Exception as ex:
                    logger.error("Failed to update created webinar information into Smartsheet. API returned error: {}".format(ex))

            else:
                # update existing event
                try:
                    w = webexApi.meetings.get(event['id'])

                    needUpdateSendEmail = \
                        event['title'] != w.title \
                        or event['startdatetime'] != datetime.fromisoformat(w.start) \
                        or event['enddatetime'] != datetime.fromisoformat(w.end)    # fromisoformat() cannot process ISO-8601 strings prior to Python 3.11, thus remove the 'Z'

                    needUpdate = \
                        needUpdateSendEmail \
                        or event['agenda'] != w.agenda

                    if needUpdate:
                        w = webexApi.meetings.update(
                            meetingId=event['id'],
                            title=event['title'],
                            agenda=event['agenda'],
                            scheduledType=event['scheduledType'],
                            start=str(event["startdatetime"]),
                            end=str(event["enddatetime"]),
                            # timezone=event['timezone'],
                            # siteUrl=event['siteUrl'],
                            password=event['password'] or w.password,    # password is required for update()
                            panelistPassword=event['panelistPassword'],
                            # reminderTime=event['reminderTime'],
                            # registration=event['registration'],
                            enabledJoinBeforeHost=event['enabledJoinBeforeHost'],
                            joinBeforeHostMinutes=event['joinBeforeHostMinutes'],                            
                            sendEmail=needUpdateSendEmail
                        )
                        logger.warning("Updated webinar information: {}".format(w.title))
                except Exception as ex:
                    logger.error("Failed to update webinar \"{}\". API returned error: {}".format(event['title'], ex))
                    try:
                        for err in ex.details['errors']:
                            logger.error("  " + err['description'])
                    except Exception:
                        pass
                    continue

                # refresh webinar registrant count in Smartsheet
                try:
                    # TODO complete when list-meeting-registrants endpoint is added to SDK
                    pass

                    registrantCount = sum(1 for _ in webexApi.meeting_invitees.list(w.id))

                    newCells = []
                    if 'registrantCount' in ssColumnMap:
                        newCells.append(ssApi.models.Cell({
                            'column_id': ssColumnMap['registrantCount'],
                            'value': registrantCount
                        }))
                    else:
                        raise Exception("No column in Smartsheet to save Registration Count.")

                    newRow = ssApi.models.Row()
                    newRow.id = ssRow.id
                    newRow.cells.extend(newCells)
                    ss = ssApi.Sheets.update_rows(ssSheet.id, [newRow])
                    logger.info("Refreshed webinar Registration Count in Smartsheet.")
                except Exception as ex:
                    logger.error("Failed to refresh webinar Registration Count in Smartsheet. API returned error: {}".format(ex))

            # update invitees (panelists and cohosts) for created or updated event
            try:
                # collect currently invited panelists and cohosts
                # also serves as an "uninvite list" - checked invitees are removed from the list
                # if there are any remaining, they will be uninvited
                currentInvitees = {}
                for i in webexApi.meeting_invitees.list(w.id, panelist=True):
                    if i.panelist or i.coHost:
                        currentInvitees[i.email] = i
            except Exception as ex:
                logger.error("Failed to process invitees for webinar \"{}\". API returned error: {}".format(event['title'], ex))
            else:

                # process panelists and cohosts

                if getWebinarProperty('noCohosts'):
                    # treat chohosts as panelists
                    event['panelists'].update(event['cohosts'])
                    event['cohosts'] = {}

                eventInvitees = event['panelists'] | event['cohosts']    # merged dicts: https://peps.python.org/pep-0584/
                for email in eventInvitees:
                    if email in currentInvitees:
                        # already invited
                        if eventInvitees[email] != currentInvitees[email].displayName \
                                or (email in event['cohosts']) != currentInvitees[email].coHost:
                            # name or status changed
                            try:
                                r = webexApi.meeting_invitees.update(
                                    meetingInviteeId=currentInvitees[email].id,
                                    email=email,
                                    displayName=eventInvitees[email],
                                    panelist=email in event['panelists'] or email in event['cohosts'],    # cohosts must also be panelists as per Webex API behavior
                                    coHost=email in event['cohosts'],
                                    sendEmail=True
                                )
                                logger.info("Updated invitee {} <{}>".format(eventInvitees[email], email))
                            except Exception as ex:
                                logger.error("Failed to update invitee \"{}\" for webinar \"{}\". API returned error: {}".format(email, event['title'], ex))
                        del currentInvitees[email]    # remove processed from the uninvite list
                    else:
                        # new, need to invite
                        try:
                            r = webexApi.meeting_invitees.create(
                                meetingId=w.id,
                                email=email,
                                displayName=eventInvitees[email],
                                panelist=email in event['panelists'] or email in event['cohosts'],    # cohosts must also be panelists as per Webex API behavior
                                coHost=email in event['cohosts'],
                                sendEmail=True
                            )
                            logger.info("Invited {} <{}>".format(eventInvitees[email], email))
                        except Exception as ex:
                            logger.error("Failed to create invitee \'{}\' for webinar \"{}\". API returned error: {}".format(email, event['title'], ex))
                # uninvite panelists/cohosts who remained in the uninvite list
                for email in currentInvitees:
                    try:
                        r = webexApi.meeting_invitees.delete(
                            meetingInviteeId=currentInvitees[email].id
                        )
                        logger.info("Uninvited {} <{}>".format(currentInvitees[email].displayName, email))
                    except Exception as ex:
                        logger.error("Failed to delete invitee \"{}\" from webinar \"{}\". API returned error: {}".format(email, event['title'], ex))

    # /for
    
    logger.warning("Done.")

    #
    # Process logs and close logging
    #
    try:
        with tempfile.NamedTemporaryFile(
            prefix=datetime.utcnow().strftime("%Y%m%d-%H%M%S "),
            suffix=".txt",
            mode="wt",
            encoding="utf-8",
            delete=False
        ) as tmp:
            tmp.write(fullLogString.getvalue())

        botApi.messages.create(
            roomId=WEBEX_BOT_ROOM_ID,
            text="Done creating and updating webinars. Full log attached. Brief log follows.\n\n" + briefLogString.getvalue(),
            files=[tmp.name]
        )

        os.remove(tmp.name)
    except Exception as ex:
        logger.error("Failed to post log into Webex bot room. " + str(ex))

    briefLogString.close()
    fullLogString.close()
