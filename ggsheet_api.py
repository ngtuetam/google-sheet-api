import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

SAMPLE_SPREADSHEET_ID = "1fC4iEOQpejG1BM-S9PlLiPRvrxQc3Pkhz8pQCZND2bc"

def authen():
  creds = None

  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=8443)
    # save
    with open("token.json", "w") as token:
      token.write(creds.to_json())
  return creds

def create_spreadsheet(title):
  creds = authen()
  try:
    service = build("sheets", "v4", credentials=creds)
    body = {
        "properties": {
          "title": title
          }
        }
    spreadsheet = (
        service.spreadsheets()
        .create(body=body)
        .execute()
    )
    print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
    return spreadsheet.get("spreadsheetId")
  except HttpError as error:
    print(f"An error occurred: {error}")
    return error
  
def create_sheet(spreadsheet_id, sheet_name):
  creds = authen()
  try:
    service = build("sheets", "v4", credentials=creds)
    body = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": sheet_name
                    }
                }
            }
        ]
    }
    result = (
        service.spreadsheets()
       .batchUpdate(spreadsheetId=spreadsheet_id, body=body)
       .execute()
    )
    print(f"Sheet ID: {result.get('replies')[0].get('addSheet').get('properties').get('sheetId')}")
    return result.get("replies")[0].get("addSheet").get("properties").get("sheetId")
  except HttpError as error:
    print(f"An error occurred: {error}")

def get_sheet_id(spreadsheet_id, sheet_name):
  creds = authen()
  try:
    service = build("sheets", "v4", credentials=creds)
    result = (
        service.spreadsheets()
       .get(spreadsheetId=spreadsheet_id)
       .execute()
    )
    sheet_id = None
    for sheet in result.get("sheets"):
      if sheet.get("properties").get("title") == sheet_name:
        sheet_id = sheet.get("properties").get("sheetId")
    print(f"Sheet ID: {sheet_id}")
    return sheet_id
  except HttpError as error:
    print(f"An error occurred: {error}")

def update_sheet(spreadsheet_id, sheet_name, range_name, body):
  creds = authen()
  try:
    service = build("sheets", "v4", credentials=creds)
    result = (
        service.spreadsheets()
       .values()
       .update(spreadsheetId=spreadsheet_id, range=range_name, valueInputOption="USER_ENTERED", body=body)
       .execute()
    )
    print(f"Update sheet: {result}")
  except HttpError as error:
    print(f"An error occurred: {error}")


if __name__ == "__main__":
  # create_spreadsheet("mysheet3")
  # create_sheet(SAMPLE_SPREADSHEET_ID, "mysheet4")
  update_sheet(SAMPLE_SPREADSHEET_ID, "mysheet4", "Sheet1!A1:f14", {"values": [["New", "Data"], ["Another", "Row"], ["New", "Row"]]})