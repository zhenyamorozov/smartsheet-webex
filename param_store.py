"""
    Utilities for interactibg with Parameter Store. Implemented for AWS SSM Parameter Store.
"""

import json
import time

import webexteamssdk

import boto3


def getSmartsheetId():
    """Returns the saved Smartsheet ID from Parameter Store
    
    Args:
        None
    Returns:
        ssSheetId: Smartsheet ID from Parameter Store
    """

    # load parameters from parameter store
    ssm_client = boto3.client("ssm")
    ssmStoredParameter = ssm_client.get_parameter(
        Name = "/daedalus/smartsheetSheetId",
        WithDecryption = True
    )
    ssSheetId = ssmStoredParameter['Parameter']['Value']
    ssm_client.close()
    return ssSheetId


def saveSmartsheetId(sheetId):
    """Saves Smartsheet ID to Parameter Store
    
    Args:
        sheetId (str)
    Returns:
        None
    """
    ssm_client = boto3.client("ssm")
    ssmStoredParameter = ssm_client.put_parameter(
        Name = "/daedalus/smartsheetSheetId",
        Value = sheetId,
        Type = "String",
        Overwrite = True
    )
    return


def getWebexIntegrationToken(webex_integration_client_id, webex_integration_client_secret):
    """Returns a fresh, usable Webex Integration access_token. 

    Webex Integration access tokens are acquired through OAuth and must be refreshed regularly.
    OAuth-provided access token and refresh token have limited lifetimes. As of now,
        access_token lifetime is 14 days since creation
        refresh_token lifetime is 90 days since last use
    This function reads tokens from Parameter Store, refreshes the access_token if it's more than halftime ols,
    and returns the access_token.

    Args:
        webex_integration_client_id - used if access token refresh is needed
        webex_integration_client_secret - used if access token refresh is needed

    Returns:
        access_token: fresh, usable Webex Integration access token
    """

    # read access tokens from Parameter Store

    ssm_client = boto3.client("ssm")
    ssmStoredParameter = ssm_client.get_parameter(
        Name = "/daedalus/webexTokens",
        WithDecryption = True
    )
    currentTokens = json.loads(ssmStoredParameter['Parameter']['Value'])
    accessToken = currentTokens['access_token']
    createdTime = currentTokens['created']
    lifetime = 14*24*60*60 # 14 days
    if createdTime + lifetime/2 < time.time():
        # refresh token
        refreshToken = currentTokens['refresh_token']

        webexApi = webexteamssdk.WebexTeamsAPI(access_token=accessToken) # passing expired access_token should still work, the API object can be initiated with any string
        newTokens = webexApi.access_tokens.refresh(
            client_id = webex_integration_client_id, 
            client_secret = webex_integration_client_secret,
            refresh_token = refreshToken
        )
        
        # save the new access tokin to the Parameter Store
        saveWebexIntegrationTokens(dict(newTokens.json_data))

    ssm_client.close()
    return accessToken


def saveWebexIntegrationTokens(tokens):
    """Saves Webex Integration tokens to Parameter Store. Adds `created` timestamp for token lifetime tracking.

    Args:
        tokens: dict of Webex Integration tokens data, as it comes from the API call

    Returns:
        None
    """
    tokens["created"] = time.time()
    ssm_client = boto3.client("ssm")
    ssmStoredParameter = ssm_client.put_parameter(
        Name = "/daedalus/webexTokens",
        Value = json.dumps(tokens),
        Type = "SecureString",
        Overwrite = True
    )
    ssm_client.close()
