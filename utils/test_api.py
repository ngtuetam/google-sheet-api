import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from _const import *


class GoogleSheetAPI:
    def __init__(self):
        self.SCOPES = SCOPES
        self.SPREADSHEET_ID = ""

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
            # save
            with open(PATH_TOKEN, "w") as token:
                token.write(creds.to_json())
        return creds

    def create_spreadsheet(self, title):
        creds = self.authenticate()
        try:
            service = build("sheets", "v4", credentials=creds)
            body = {"properties": {"title": title}}
            spreadsheet = service.spreadsheets().create(body=body).execute()
            print(f"Spreadsheet ID: {spreadsheet.get('spreadsheetId')}")
            return spreadsheet.get("spreadsheetId")
        except HttpError as error:
            print(f"Error: {error}")
            return error

    def create_sheet(self, spreadsheet_id, sheet_name):
        creds = self.authenticate()
        try:
            service = build("sheets", "v4", credentials=creds)
            body = {
                "requests": [
                    {
                        "addSheet": {
                            "properties": {"title": sheet_name}
                        }
                    }
                ]
            }
            result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
            sheet_id = result.get("replies")[0].get("addSheet").get("properties").get("sheetId")
            print(f"Sheet ID: {sheet_id}")
            return sheet_id
        except HttpError as error:
            print(f"Error: {error}")

    def get_sheet_id(self, spreadsheet_id, sheet_name):
        creds = self.authenticate()
        try:
            service = build("sheets", "v4", credentials=creds)
            result = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_id = None
            for sheet in result.get("sheets"):
                if sheet.get("properties").get("title") == sheet_name:
                    sheet_id = sheet.get("properties").get("sheetId")
            print(f"Sheet ID: {sheet_id}")
            return sheet_id
        except HttpError as error:
            print(f"Error: {error}")

    def update_sheet(self, spreadsheet_id, sheet_name, range_name, body):
        creds = self.authenticate()
        try:
            service = build("sheets", "v4", credentials=creds)
            result = (
                service.spreadsheets()
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


class ExportGoogleSheet:
    def __init__(self, google_sheet_api):
        self.google_sheet_api = google_sheet_api

    def export_table_to_sheet(self, spreadsheet_id, sheet_name, table_data):
        creds = self.google_sheet_api.authenticate()

        try:
            service = build("sheets", "v4", credentials=creds)
            sheet_id = self.create_sheet(service, spreadsheet_id, sheet_name)
            self.populate_sheet(service, spreadsheet_id, sheet_id, table_data, sheet_name)
            print(f"Table data exported to sheet '{sheet_name}' in spreadsheet ID: {spreadsheet_id}")
        except HttpError as error:
            print(f"Error: {error}")

    def create_spreadsheet_and_import_table(self, spreadsheet_title, sheet_name, table_data):
        creds = self.google_sheet_api.authenticate()

        try:
            service = build("sheets", "v4", credentials=creds)
            spreadsheet_id = self.google_sheet_api.create_spreadsheet(spreadsheet_title)
            sheet_id = self.google_sheet_api.create_sheet(service, spreadsheet_id, sheet_name)
            self.populate_sheet(service, spreadsheet_id, sheet_id, table_data)
            print(f"Table data imported to sheet '{sheet_name}' in the new spreadsheet ID: {spreadsheet_id}")
            return spreadsheet_id, sheet_id
        except HttpError as error:
            print(f"Error: {error}")
            return None, None

    def update_existing_sheet(self, spreadsheet_id, sheet_name, table_data, sheet_range):
        creds = self.google_sheet_api.authenticate()

        try:
            service = build("sheets", "v4", credentials=creds)
            sheet_id = self.google_sheet_api.get_sheet_id(spreadsheet_id, sheet_name)

            if sheet_id is not None:
                range_name = f"{sheet_name}!{sheet_range}"
                body = {"values": [list(table_data.keys())] + list(zip(*table_data.values()))}
                self.google_sheet_api.update_sheet(spreadsheet_id, sheet_id, range_name, body)
                print(f"Table data updated in sheet '{sheet_name}' of spreadsheet ID: {spreadsheet_id}")
            else:
                print(f"Sheet '{sheet_name}' not found in spreadsheet ID: {spreadsheet_id}")
        except HttpError as error:
            print(f"Error: {error}")

    def create_sheet(self, service, spreadsheet_id, sheet_name):
        return self.google_sheet_api.create_sheet(spreadsheet_id, sheet_name)

    def populate_sheet(self, service, spreadsheet_id, sheet_id, table_data, sheet_name):
        range_name = f"{sheet_name}!A1"
        body = {"values": [list(table_data.keys())] + list(zip(*table_data.values()))}
        self.google_sheet_api.update_sheet(spreadsheet_id, sheet_id, range_name, body)
