
from googlesheet_api_handle import HandleGoogleSheetAPI
from utils.test_api import GoogleSheetAPI, ExportGoogleSheet

spreadsheet_id = "1IYgieoLhDpmzDc6riOXvCj2KAs1fggvgUOyQC5Jxirg"
handle_google_sheet_api = HandleGoogleSheetAPI()

# TEST 1
spreadsheet_title = "New Spreadsheet NEW"
# sheet_name = "NewSheet"
table_data = {
    "Column1": [7, 8, 9],
    "Column2": ["P", "Q", "R"],

}
spreadsheet_id, sheet_id = handle_google_sheet_api.create_spreadsheet_and_import_table(spreadsheet_title, table_data)

# TEST 2
new_sheet_name = "NewSheet2"
table_data_new_sheet = {
    "Column1": [10, 11, 12],
    "Column2": ["X", "Y", "Z"],

}
new_sheet_id = handle_google_sheet_api.create_new_sheet_and_import_table(spreadsheet_id, new_sheet_name, table_data_new_sheet)

# TEST 3
sheet_name = "NewSheet2"
table_data_existing_sheet = {
    "Column1": [13, 14, 15,16],
    "Column2": ["Tam", "Test", "API","GG api"],
   
}
sheet_range = "A1:B5"
handle_google_sheet_api.update_existing_sheet(spreadsheet_id,sheet_name, table_data_existing_sheet,sheet_range)


# google_sheet_api = GoogleSheetAPI()
# spreadsheet_id = google_sheet_api.create_spreadsheet("My Spreadsheet")
# sheet_id = google_sheet_api.create_sheet(spreadsheet_id, "Sheet1")
# google_sheet_api.get_sheet_id(spreadsheet_id, "Sheet1")


# google_sheet_api = GoogleSheetAPI()
# table_exporter = ExportGoogleSheet(google_sheet_api)