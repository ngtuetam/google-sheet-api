import os.path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from googlesheet_api_auth import Authenticator
from _const import *


class GoogleSheetAPI(Authenticator):
    SCOPES = SCOPES

    def __init__(self):
        super().__init__()
        self.creds = self.authenticate()
        self.service = build("sheets", "v4", credentials=self.creds)

    def create_spreadsheet(self, title):
        try:
            body = {
                    "properties": {
                        "title": title
                        }
                    }
            spreadsheet = self.service.spreadsheets().create(body=body).execute()
            print(f"Spreadsheet ID: {spreadsheet.get('spreadsheetId')}")
            return spreadsheet.get("spreadsheetId")
        except HttpError as error:
            print(f"Error: {error}")
            return None

    def create_sheet(self, spreadsheet_id, sheet_name):
        try:
            body = {
                "requests": [
                    {
                        "addSheet": {
                            "properties": {"title": sheet_name}
                        }
                    }
                ]
            }
            result = self.service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
            sheet_id = result.get("replies")[0].get("addSheet").get("properties").get("sheetId")
            print(f"Sheet ID: {sheet_id}")
            return sheet_id
        except HttpError as error:
            print(f"Error: {error}")
            return None

    def get_sheet_id(self, spreadsheet_id, sheet_name):
        try:
            result = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None
            for sheet in result.get("sheets"):
                if sheet.get("properties").get("title") == sheet_name:
                    sheet_id = sheet.get("properties").get("sheetId")
            print(f"Sheet ID: {sheet_id}")
            return sheet_id
        except HttpError as error:
            print(f"Error: {error}")
            return None

    def update_sheet(self, spreadsheet_id, sheet_name, range_name, body):
        try:
            result = (
                self.service.spreadsheets()
                .values()
                .update(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueInputOption="USER_ENTERED",
                    body=body,
                )
                .execute()
            )
            print(f"Update sheet: {result}")
        except HttpError as error:
            print(f"Error: {error}")


