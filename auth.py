from __main__ import *

from flask import Flask, redirect, request, session

from param_store import *

import os
from uuid import uuid4
import urllib.parse
import requests


# OAuth static vars
oa_authorizationURI = "https://api.ciscospark.com/v1/authorize?"
oa_tokenURI = "https://api.ciscospark.com/v1/access_token"
if os.getenv("FLASK_ENV") == "development":
    # dev
    oa_callbackUri = "http://localhost:5000"+"/callback"
else:
    # prod
    oa_callbackUri = webAppPublicUrl+"/callback"


@app.route("/")
def root():
    print("/ requested")
    return "Hey, this is the bot running on Flask!"

# OAuth Step 1
# this function is never to be used as the auth URL is provided directly
@app.route("/auth")
def auth():
    """Step 1: User Authorization.
    Redirect the user/resource owner to the OAuth provider
    using an URL with a few key OAuth parameters.
    """

    state = str(uuid4())
    # print("state: ", state)
    # # State is used to prevent CSRF, save state in session
    # session['oauth_state'] = state
    # session.modified = True

    oa_params = {'response_type':"code",
              'client_id': WEBEX_INTEGRATION_CLIENT_ID,
              'redirect_uri': oa_callbackUri,
              'scope':"meeting:schedules_read meeting:schedules_write spark:all meeting:preferences_read meeting:recordings_read meeting:participants_read",
              'state': state
              }
    oa_authorizationFullURI = oa_authorizationURI+urllib.parse.urlencode(oa_params)
    # print(oa_authorizationFullURI)


    return redirect(oa_authorizationFullURI)

# OAuth Step 2
# Step 2: User authorization, this happens on the provider side.

# OAuth Step 3
@app.route("/callback", methods=["GET"])
def callback():
    """ Step 3: Retrieving an access token.
    The user has been redirected back from the provider to your registered
    callback URL. With this redirection comes an authorization code included
    in the redirect URL. We will use that to obtain an access token.
    """
    print("OAuth callback start")
    
    oa_error = request.args.get("error", '')
    if oa_error:
        return "Error: " + oa_error

    oa_code = request.args.get("code")
    if not oa_code:
        return "Authorization failed. Authorization provider did not return authorization code."

    # oa_state = request.args.get('state', '')
    # if not oa_state:
    #     return "Authorization failed. Authorization provider did not return state."
    # if oa_state != session['oauth_state']:
    #     return "Authorization failed. State does not match."

    oa_data = {
        'grant_type': "authorization_code",
        'redirect_uri': oa_callbackUri,
        'code': oa_code,
        'client_id': WEBEX_INTEGRATION_CLIENT_ID,
        'client_secret': WEBEX_INTEGRATION_CLIENT_SECRET
    }

    oa_resp = requests.post(oa_tokenURI, data=oa_data)
    oa_tokens = oa_resp.json()
    # print(oa_tokens)
    if not oa_resp.ok:
        return "Authorization failed. Authorization provider returned:<br> " + str(oa_tokens)
    else:
        try:
            saveWebexIntegrationTokens(oa_tokens)
        except Exception as ex:
            return "Webex authorization was successful, but could not save tokens to Parameter Store. Check local AWS configuration."
        return "Authorization successful. Access and refresh tokens retrieved and saved in Parameter Store."
