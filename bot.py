from __main__ import *

from flask import request, url_for

import schedule

from param_store import (
    getSmartsheetId,
    saveSmartsheetId,
    getWebexIntegrationToken
)

import os
import urllib.parse
from datetime import datetime

import webexteamssdk

from webexteamssdk import WebexTeamsAPI
from webexteamssdk.models.cards.card import AdaptiveCard
from webexteamssdk.models.cards.inputs import *
from webexteamssdk.models.cards.components import *
from webexteamssdk.models.cards.container import *
from webexteamssdk.models.cards.actions import *
from webexteamssdk.models.cards.options import *

import smartsheet

webhookTargetUrl = webAppPublicUrl + "/webhook"

# initialize Webex Teams bot control object
try:
    botApi = WebexTeamsAPI(WEBEX_BOT_TOKEN)    # this will not raise an exception, even if bot token isn't correct, so,
    # need to make an API call to check API is functional
    assert botApi.people.me()
except Exception:
    print("Could not initialize Webex bot API object.")
    raise SystemExit()


# print("Creating webhooks in Webex")

# delete ALL current webhooks - this bot is supposed to be used only with one instance of the app
for wh in botApi.webhooks.list():
    botApi.webhooks.delete(wh.id)

# create new webhooks
try:
    botApi.webhooks.create(
        name="Smartsheet-Webex bot - attachmentActions",
        targetUrl=webhookTargetUrl,
        resource="attachmentActions",
        event="created",
        filter="roomId=" + WEBEX_BOT_ROOM_ID
    )
except Exception:
    print("Could not create a Webex bot API webhook.")
try:
    botApi.webhooks.create(
        name="Smartsheet-Webex bot - messages",
        targetUrl=webhookTargetUrl,
        resource="messages",
        event="created",
        filter="roomId=" + WEBEX_BOT_ROOM_ID
    )
except Exception:
    print("Could not create a Webex bot API webhook.")


@application.route("/webhook", methods=['GET', 'POST'])
def webhook():
    # print ("Webhook arrived.")
    # print(request)

    webhookJson = request.json

    # check if the received webhook is properly formed and relevant
    try:
        assert webhookJson['resource'] in ("messages", "attachmentActions")
        assert webhookJson['event'] == "created"
        assert webhookJson['data']['roomId'] == WEBEX_BOT_ROOM_ID
    except Exception:
        print("The arrived webhook is malformed or does not indicate an actionable event in the log and control room")
        return "Webhook processed."

    # will need our own name
    me = botApi.people.me()

    # static adaptive card - greeting
    greetingCard = AdaptiveCard(
        fallbackText="Hi, I am {}, I automatically create Webex Webinar sessions based on information in a Smartsheet. Adaptive cards feature is required to use me.".format(me.nickName),
        body=[
            TextBlock(
                text="Smartsheet to Webex Webinar automation",
                weight=FontWeight.BOLDER,
                size=FontSize.MEDIUM,
            ),
            TextBlock(
                text="Hi, I am {}, I automatically create Webex Webinar sessions based on information in a Smartsheet.".format(me.nickName),
                wrap=True,
            )

        ],
        actions=[
            Submit(title="Schedule now", data={'act': "schedule now"}),
            Submit(title="Set Smartsheet", data={'act': "set smartsheet"}),
            Submit(title="Authorize Webex", data={'act': "authorize webex"}),
            Submit(title="?", data={'act': "help"}),
        ]
    )

    # if webhook indicates a message sent by us (the bot itself), ignore it
    if webhookJson['data']['personId'] == me.id:
        return "Webhook processed."

    # received a text message
    if webhookJson['resource'] == "messages":
        # retrieve the new message details
        # message = botApi.messages.get(webhookJson['data']['id'])
        # print(message)

        # respond with the greeting card to any message
        botApi.messages.create(text=greetingCard.fallbackText, roomId=WEBEX_BOT_ROOM_ID, attachments=[greetingCard])

    # received a card action
    elif webhookJson['resource'] == "attachmentActions":

        # retrieve the new attachment action details
        action = botApi.attachment_actions.get(webhookJson['data']['id'])
        # print("Action:\n", action)

        # print("actionInputs", actionInputs)

        # "?" (help) action
        if action.type == "submit" and action.inputs['act'] == "help":
            botApi.messages.create(markdown="""
Smartsheet and Webex Automation creates webinars in Webex Webinar based on information in Smartsheet.
It is easy to use:
1. Collaborate with your team on webinar planning in Smartsheet. When ready for creation, check **Create=yes**
2. Click **Schedule Now** button to start webinar scheduling process
3. Webinars are created

Features and basic usage: https://github.com/zhenyamorozov/smartsheet-webex#smartsheet-and-webex-automation
How to set up and get started: https://github.com/zhenyamorozov/smartsheet-webex/blob/master/docs/get_started.rst#get-started
            """, roomId=WEBEX_BOT_ROOM_ID)
            # resend greeting card
            botApi.messages.create(text=greetingCard.fallbackText, roomId=WEBEX_BOT_ROOM_ID, attachments=[greetingCard])
            pass

        # "Schedule now" action
        if action.type == "submit" and action.inputs['act'] == "schedule now":
            botApi.messages.create(text="The process to create/update webinars has started.", roomId=WEBEX_BOT_ROOM_ID)
            # invoke the webinar scheduling process
            schedule.run()

        # "Set Smartsheet" action
        if action.type == "submit" and action.inputs['act'] == "set smartsheet":
            try:
                # load Smartsheet ID from parameter store
                sheetId = getSmartsheetId()

                # fetch current smartsheet
                ssApi = smartsheet.Smartsheet()
                ssApi.errors_as_exceptions(True)

                sheetName = ssApi.Sheets.get_sheet(sheetId).name

            except Exception:
                sheetId = ""
                sheetName = ""

            card = AdaptiveCard(
                fallbackText="Adaptive cards feature is required to use me.",
                body=[
                    TextBlock(
                        text="Smartsheet setting",
                        weight=FontWeight.BOLDER,
                        size=FontSize.MEDIUM,
                    ),
                    TextBlock(
                        text="Information for scheduled sessions is taken from a Smartsheet. This is the current working smartsheet. You can change it here.",
                        wrap=True,
                    ),
                    FactSet(
                        facts=[
                            Fact(
                                title="Name",
                                value=str(sheetName)
                            ),
                            Fact(
                                title="ID",
                                value=sheetId
                            )
                        ]
                    )
                ],
                actions=[
                    ShowCard(
                        title="Change",
                        card=AdaptiveCard(
                            body=[
                                TextBlock(
                                    text="Smartsheet ID can be found  under File - Properties. It is a numeric value. Alternatively, you can just copy the Smartsheet URL here. ",
                                    wrap=True,
                                ),
                                Text('newSmartsheetId', placeholder="New Smartsheet ID", isMultiline=False),
                            ],
                            actions=[
                                Submit(title="OK", data={'act': "save smartsheet id"}),
                            ]
                        )
                    ),
                    Submit(title="Create Template", data={'act': "create smartsheet template"})
                ]
            )

            # print(card.to_json())
            botApi.messages.create(text="Could not send the action card", roomId=WEBEX_BOT_ROOM_ID, attachments=[card])
            pass

        # "Save Smartsheet ID" action
        if action.type == "submit" and action.inputs['act'] == "save smartsheet id":
            # print(action)

            if 'newSmartsheetId' not in action.inputs or not action.inputs['newSmartsheetId'].strip():
                botApi.messages.create(
                    text="Smartsheet ID cannot be empty.",
                    roomId=WEBEX_BOT_ROOM_ID
                )
            else:
                try:
                    newSheetId = action.inputs['newSmartsheetId'].strip()

                    # if numeric sheet ID provided, keep it
                    if newSheetId.isdigit():
                        pass
                    # if sheet URL is provided, use the last part of the path
                    elif newSheetId.startswith("http"):
                        newSheetId = urllib.parse.urlparse(newSheetId).path.split('/')[-1]
                    # else try to use the provided ID as-is
                    else:
                        pass

                    # fetch the new smartsheet
                    ssApi = smartsheet.Smartsheet()
                    ssApi.errors_as_exceptions(True)

                    # this will fetch the new sheet by numeric ID or the ID from the URL
                    newSheet = ssApi.Sheets.get_sheet(newSheetId)

                    newSheetName = newSheet.name
                    newSheetId = newSheet.id

                    try:
                        saveSmartsheetId(str(newSheetId))
                        # send cpnfirmation message
                        botApi.messages.create(
                            markdown="New Smartsheet is set:\n``{}``".format(newSheetName),
                            roomId=WEBEX_BOT_ROOM_ID
                        )
                        # resend greeting card
                        botApi.messages.create(text=greetingCard.fallbackText, roomId=WEBEX_BOT_ROOM_ID, attachments=[greetingCard])
                    except Exception:
                        botApi.messages.create(
                            text="Could not save new Smartsheet ID to Parameter Store. Check local AWS configuration.",
                            roomId=WEBEX_BOT_ROOM_ID
                        )
                except Exception:
                    botApi.messages.create(
                        text="That Sheet ID did not work. Try again.",
                        roomId=WEBEX_BOT_ROOM_ID
                    )

        # "Create Smartsheet Template" action
        if action.type == "submit" and action.inputs['act'] == "create smartsheet template":
            # print(action)
            try:
                ssApi = smartsheet.Smartsheet()
                ssApi.errors_as_exceptions(True)
                sheetSpec = smartsheet.models.Sheet({
                    'name': "Template " + datetime.utcnow().strftime("%Y%m%d-%H%M%S"),
                    'columns': [
                        {
                            'title': "Create",
                            'type': smartsheet.models.enums.column_type.ColumnType.PICKLIST,
                            'options': ["yes", "no"],
                            'description': "To check out a webinar for creation, change this value to 'yes'. Required field."
                        },
                        {
                            'title': "Start Date",
                            'type': smartsheet.models.enums.column_type.ColumnType.DATE,
                            'description': "You can change the date format in Profile icon - Personal Seetings - Settings - Regional Preferences. Required field."
                        },
                        {
                            'title': "Start Time",
                            'type': smartsheet.models.enums.column_type.ColumnType.TEXT_NUMBER,
                            'description': "24-hour clock HH:MM format. Required field."
                        },
                        {
                            'title': "Duration",
                            'type': smartsheet.models.enums.column_type.ColumnType.TEXT_NUMBER,
                            'description': "In minutes. If not specified, the standard duration is used."
                        },
                        {
                            'title': "Title",
                            'type': smartsheet.models.enums.column_type.ColumnType.TEXT_NUMBER,
                            'description': "128 characters maximum. Required field."
                        },
                        {
                            'title': "Agenda",
                            'type': smartsheet.models.enums.column_type.ColumnType.TEXT_NUMBER,
                            'description': "1300 characters maximum."
                        },
                        {
                            'title': "Cohosts",
                            'type': smartsheet.models.enums.column_type.ColumnType.MULTI_CONTACT_LIST,
                            'description': "Multiple contacts may be selected."
                        },
                        {
                            'title': "Panelists",
                            'type': smartsheet.models.enums.column_type.ColumnType.MULTI_CONTACT_LIST,
                            'description': "Comma-separated list of 'name <email>'. Nicknames can be used."
                        },
                        {
                            'primary': True,
                            'title': "Webinar ID",
                            'type': smartsheet.models.enums.column_type.ColumnType.TEXT_NUMBER,
                            'description': "Automatically populated and used for the automation. Required field."

                        },
                        {
                            'title': "Attendee URL",
                            'type': smartsheet.models.enums.column_type.ColumnType.TEXT_NUMBER,
                            'description': "Automatically populated. This is the Join URL, NOT the Registration URL."
                        },
                        {
                            'title': "Host Key",
                            'type': smartsheet.models.enums.column_type.ColumnType.TEXT_NUMBER,
                            'description': "Automatically populated."
                        },
                        {
                            'title': "Registrant Count",
                            'type': smartsheet.models.enums.column_type.ColumnType.TEXT_NUMBER,
                            'description': "Automatically populated."
                        }
                    ]
                })
                newSheet = ssApi.Home.create_sheet(sheetSpec).result
                # additional settings for columns which can't be set at sheet creation (no idea why)
                for col in newSheet.columns:
                    if col.title == "Create":
                        col_id = col.id_
                        col.id_ = col.version = None
                        col.validation = True
                        col.format = ",,,,,,,,,18,,,,,,"    # Smartsheet formatting is sorcery
                        ssApi.Sheets.update_column(newSheet.id_, col_id, col)
                    if col.title == "Start Date":
                        col_id = col.id_
                        col.id_ = col.version = None
                        col.validation = True
                        ssApi.Sheets.update_column(newSheet.id_, col_id, col)
                    if col.title in ("Webinar ID", "Attendee URL", "Host Key", "Registrant Count"):
                        col_id = col.id_
                        col.id_ = col.version = col.validation = col.primary = None
                        col.locked = True
                        col.format = ",,,,,,,,,18,,,,,,"
                        ssApi.Sheets.update_column(newSheet.id_, col_id, col)

                botApi.messages.create(
                    text="Here is your newly created Smartsheet template. Don't forget to set it as the current working smartsheet.\n{}".format(newSheet.permalink),
                    roomId=WEBEX_BOT_ROOM_ID
                )

            except Exception:
                botApi.messages.create(
                    text="Couldn't create a Smartsheet template.",
                    roomId=WEBEX_BOT_ROOM_ID
                )

        # "Authorize Webex" action
        if action.type == "submit" and action.inputs['act'] == "authorize webex":
            try:
                # get a fresh Webex Integration access token
                access_token = getWebexIntegrationToken(
                    webex_integration_client_id=WEBEX_INTEGRATION_CLIENT_ID,
                    webex_integration_client_secret=WEBEX_INTEGRATION_CLIENT_SECRET
                )

                # get information about the current authorized Webex user
                webexApi = webexteamssdk.WebexTeamsAPI(access_token)
                webexMe = webexApi.people.me()
                webexEmail = webexMe.emails[0]
                webexDisplayName = webexMe.displayName
            except Exception as ex:
                # haven't been authorized yet
                webexEmail = ""
                webexDisplayName = "Not authorized yet"

            if os.getenv("FLASK_ENV") == "development":
                # dev
                authUrl = "http://localhost:5000" + "/auth"
            else:
                # prod
                authUrl = url_for("auth", _external=True)

            card = AdaptiveCard(
                fallbackText="Adaptive cards feature is required to use me.",
                body=[
                    TextBlock(
                        text="Webex integration authorization",
                        weight=FontWeight.BOLDER,
                        size=FontSize.MEDIUM,
                    ),
                    TextBlock(
                        text="Webex integration is used to create Webex Webinar sessions. This is the user currently authorized to create sessions. You can update authorization here.",
                        wrap=True,
                    ),
                    FactSet(
                        facts=[
                            Fact(
                                title="Name",
                                value=webexDisplayName
                            ),
                            Fact(
                                title="Email",
                                value=webexEmail
                            )
                        ]
                    )
                ],
                actions=[
                    ShowCard(
                        title="Change",
                        card=AdaptiveCard(
                            body=[
                                TextBlock(
                                    text="Click the button below and complete authorization process in your broswer.",
                                    wrap=True,
                                ),
                            ],
                            actions=[
                                OpenUrl(title="Authorize", url=authUrl),
                                ShowCard(
                                    title="Copy URL",
                                    card=AdaptiveCard(
                                        body=[
                                            TextBlock(
                                                text=authUrl,
                                                wrap=True,
                                            ),
                                        ],
                                    )
                                )
                            ]
                        )
                    )
                ]
            )    # /AdaptiveCard

            botApi.messages.create(text="Could not send the action card", roomId=WEBEX_BOT_ROOM_ID, attachments=[card])

    return "webhook accepted"
