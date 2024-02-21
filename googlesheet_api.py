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

    # write data to worksheet
    def write_data_range(self, spreadsheet_id, range_name, data):
        try:
            body = {"values": data}
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
            print(f"{result.get('updatedCells')} cells updated.")
            return result
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error
        
    def format_header(self, spreadsheet_id,  cell_format=None, format_range={}):
        try:

            request = {
                "repeatCell": {
                    "range": format_range,
                    "cell": {
                        "userEnteredFormat": cell_format
                    },
                    "fields": "userEnteredFormat"
                }
            }

            response = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={'requests': [request]}).execute()

            print("Cell formatted successfully.")
            return response
        except HttpError as error:
            print(f"An error occurred: {error}")
            return None
        
    def merge_cells(self, spreadsheet_id, merge_range, merge_type='MERGE_ROWS'):
        """
            Merge cell
        Args:
            spreadsheet_id (str): spreadsheet id
            merge_range (dict): dictionary range to merge cell
        """
        try:
            requests = {
                "mergeCells": {
                    "range": merge_range,
                    "mergeType": merge_type
                }
            },
            response = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={'requests': [requests]}).execute()
            return None
        except HttpError as error:
            print(f"An error occurred: {error}")
            return None

    def merge_cell_and_write_data(self, spreadsheet_id, merge_range,  range_name, data, merge_type=None):
        for row in data:
            for i in range(0, len(row)):
                if i == merge_range['startColumnIndex']:
                    for j in range(2, merge_range['endColumnIndex']):
                        row[i] = row[i] + ' ' + row[j]
                    break
        self.merge_cells(spreadsheet_id, merge_range)
        self.write_data_range(spreadsheet_id, range_name, data)


