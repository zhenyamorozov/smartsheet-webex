"""
This web application serves two purposes:
    - process OAuth requests for Webex Integration
    - respond to Webex bot webhooks
"""

from flask import Flask, redirect, request, session

import os
from dotenv import load_dotenv
import logging
import requests

# initialize logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# load env variables
load_dotenv(override=True)

# load env vars for Flask
FLASK_SECRET_KEY = os.getenv("FLASK_SECRET_KEY", "dev")
if FLASK_SECRET_KEY=="dev": 
    logger.debug("  Flask secret key is not set in env. Set to any random string with FLASK_SECRET_KEY environment variable.")

# load env vars for Webex integration
WEBEX_INTEGRATION_CLIENT_ID = os.getenv("WEBEX_INTEGRATION_CLIENT_ID")
if not WEBEX_INTEGRATION_CLIENT_ID: 
    logger.fatal("  Webex Integration Client ID is missing. Provide with WEBEX_INTEGRATION_CLIENT_ID environment variable.")
    raise SystemExit()
WEBEX_INTEGRATION_CLIENT_SECRET = os.getenv("WEBEX_INTEGRATION_CLIENT_SECRET")
if not WEBEX_INTEGRATION_CLIENT_SECRET: 
    logger.fatal("  Webex Integration Client ID is missing. Provide with WEBEX_INTEGRATION_CLIENT_SECRET environment variable.")
    raise SystemExit()

# load env vars for Webex bot
WEBEX_BOT_TOKEN = os.getenv("WEBEX_BOT_TOKEN")
if not WEBEX_BOT_TOKEN: 
    logger.fatal("  Webex bot access token is missing. Provide with WEBEX_BOT_TOKEN environment variable.")
    raise SystemExit()
WEBEX_BOT_ROOM_ID = os.getenv("WEBEX_BOT_ROOM_ID")
if not WEBEX_BOT_ROOM_ID: 
    logger.fatal("  Webex bot room ID is missing. Provide with WEBEX_BOT_ROOM_ID environment variable.")
    raise SystemExit()

# Get the public app URL
# try to get the ngrok URL (dev)
try:
    r = requests.get("http://localhost:4040/api/tunnels")
    webAppPublicUrl = r.json()['tunnels'][0]['public_url']
except:
    # do nothing
    pass
# try to get the AWS URL (prod)
try:
    r = requests.get("http://169.254.169.254/latest/meta-data/public-hostname")
    webAppPublicUrl = r.text
except:
    # do nothing
    pass

if not webAppPublicUrl:
    logger.fatal("Could not get the web app public URL")
    raise SystemExit()

logger.debug("Launch Flask app")

app = Flask(__name__)

app.secret_key = FLASK_SECRET_KEY


import auth

import bot


if __name__ == "__main__":
    app.run()
