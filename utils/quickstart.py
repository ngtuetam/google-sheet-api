import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from google.oauth2 import service_account


SERVICE_ACCOUNT_FILE = 'keys.json'
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

SAMPLE_SPREADSHEET_ID = "1fC4iEOQpejG1BM-S9PlLiPRvrxQc3Pkhz8pQCZND2bc"
SAMPLE_RANGE_NAME = "Sheet1!A1:f14"

service = build("sheets", "v4", credentials=creds)
# call the Sheets API
sheet = service.spreadsheets()

# result = (
#     sheet.values()
#     .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
#     .execute()
# )
# values = result.get("values", [])
# print(result)

# update
body = [["New", "Data"], ["Another", "Row"], ["New", "Row"]]

request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="Sheet2!B2", valueInputOption="USER_ENTERED", body={"values": body}).execute()
print(request)