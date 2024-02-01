
from google.oauth2 import service_account
from googleapiclient.discovery import build
from ._const import *


creds = service_account.Credentials.from_service_account_file(JSON_CRED_PATH_GOOGLE_SHEET)
service = build('sheets', 'v4', credentials=creds)

spreadsheet_id = 'ID file Google sheets'
sheet_name = 'Tên sheet'
range_name = f'{sheet_name}!A1:B5'


result = service.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id, range=range_name).execute()


for row in result.get('values', []):
    print(row)