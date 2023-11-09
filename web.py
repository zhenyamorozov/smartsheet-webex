"""
This web application serves two purposes:
    - process OAuth requests for Webex Integration
    - respond to Webex bot webhooks
"""

from flask import Flask

import os
from dotenv import load_dotenv
import requests

# load env variables
load_dotenv(override=True)

# load env vars for Flask
FLASK_SECRET_KEY = os.getenv('FLASK_SECRET_KEY', "dev")
if FLASK_SECRET_KEY == "dev":
    print("Flask secret key is not set in env. Set to any random string with FLASK_SECRET_KEY environment variable.")

# load env vars for Webex integration
WEBEX_INTEGRATION_CLIENT_ID = os.getenv('WEBEX_INTEGRATION_CLIENT_ID')
if not WEBEX_INTEGRATION_CLIENT_ID:
    print("Webex Integration Client ID is missing. Provide with WEBEX_INTEGRATION_CLIENT_ID environment variable.")
    raise SystemExit()
WEBEX_INTEGRATION_CLIENT_SECRET = os.getenv('WEBEX_INTEGRATION_CLIENT_SECRET')
if not WEBEX_INTEGRATION_CLIENT_SECRET:
    print("Webex Integration Client Secret is missing. Provide with WEBEX_INTEGRATION_CLIENT_SECRET environment variable.")
    raise SystemExit()

# load env vars for Webex bot
WEBEX_BOT_TOKEN = os.getenv('WEBEX_BOT_TOKEN')
if not WEBEX_BOT_TOKEN:
    print("Webex bot access token is missing. Provide with WEBEX_BOT_TOKEN environment variable.")
    raise SystemExit()
WEBEX_BOT_ROOM_ID = os.getenv('WEBEX_BOT_ROOM_ID')
if not WEBEX_BOT_ROOM_ID:
    print("Webex bot room ID is missing. Provide with WEBEX_BOT_ROOM_ID environment variable.")
    raise SystemExit()

# Get the public app URL
webAppPublicUrl = ""
# try to get the ngrok URL (dev)
try:
    r = requests.get("http://localhost:4040/api/tunnels")
    webAppPublicUrl = r.json()['tunnels'][0]['public_url']
except Exception:
    # do nothing
    pass
# try to get the AWS URL (prod)
try:
    # AWS IMDSv2 requires authenticatiom
    r = requests.put("http://169.254.169.254/latest/api/token", 
        headers={"X-aws-ec2-metadata-token-ttl-seconds": "21600"})
    imdsToken = r.text
    r = requests.get("http://169.254.169.254/latest/meta-data/public-hostname", 
        headers={"X-aws-ec2-metadata-token": imdsToken})
    webAppPublicUrl = "https://" + r.text
except Exception:
    # do nothing
    pass

if not webAppPublicUrl:
    print("Could not get the web app public URL")
    raise SystemExit()

print("Launch Flask app")

application = Flask(__name__)

application.secret_key = FLASK_SECRET_KEY


import auth

import bot


if __name__ == "__main__":
    application.run()
