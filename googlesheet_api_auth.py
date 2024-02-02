import os.path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from _const import *


class Authenticator:
    def authenticate(self):
        creds = None

        if os.path.exists(PATH_TOKEN):
            creds = Credentials.from_authorized_user_file(PATH_TOKEN, self.SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(PATH_CRED_GG_SHEET, self.SCOPES)
                creds = flow.run_local_server(port=8443)
            with open(PATH_TOKEN, "w") as token:
                token.write(creds.to_json())
        return creds
