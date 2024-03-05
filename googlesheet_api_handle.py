from googlesheet_api import GoogleSheetAPI
from _const import *


class ExportGoogleSheet(GoogleSheetAPI):
    def __init__(self):
        super().__init__()

    def create_spreadsheet_and_import_table(self, spreadsheet_title,table_data,sheet_name=None):
        spreadsheet_id = self.create_spreadsheet(spreadsheet_title)
        sheet_id = None
        if spreadsheet_id:
            if sheet_name is not None:
                sheet_id = self.create_sheet(spreadsheet_id, sheet_name)
            else:
                sheet_name = "Sheet1"
            if sheet_id or sheet_name:
                self.write_sheet(spreadsheet_id, sheet_id, sheet_name, table_data, sheet_range="A1")
                print(f"Table data imported to sheet '{sheet_name}' in the new spreadsheet ID: {spreadsheet_id}")
                return spreadsheet_id, sheet_id
            else:
                print("Failed to create sheet.")
        else:
            print("Failed to create spreadsheet.")
        return None, None

    def create_new_sheet_and_import_table(self, spreadsheet_id, sheet_name, table_data, sheet_range=None):
        sheet_id = self.create_sheet(spreadsheet_id, sheet_name)
        if sheet_range is None:
            sheet_range = "A1"
        if sheet_id:
            self.write_sheet(spreadsheet_id, sheet_id, sheet_name, table_data, sheet_range)
            print(f"Table data imported to new sheet '{sheet_name}' in spreadsheet ID: {spreadsheet_id}")
            return sheet_id
        else:
            print("Failed to create sheet.")
        return None

    def update_existing_sheet(self, spreadsheet_id, sheet_name, table_data, sheet_range):
        sheet_id = self.get_sheet_id(spreadsheet_id, sheet_name)
        if sheet_id:
            range_name = f"{sheet_name}!{sheet_range}"
            body = {"values": [list(table_data.keys())] + list(zip(*table_data.values()))}
            self.update_sheet(spreadsheet_id, sheet_id, range_name, body)
            print(f"Table data updated in sheet '{sheet_name}' of spreadsheet ID: {spreadsheet_id}")
        else:
            print(f"Sheet '{sheet_name}' not found in spreadsheet ID: {spreadsheet_id}")

    def write_sheet(self, spreadsheet_id, sheet_id, sheet_name, table_data, sheet_range):
        range_name = f"{sheet_name}!{sheet_range}"
        body = {"values": [list(table_data.keys())] + list(zip(*table_data.values()))}
        self.update_sheet(spreadsheet_id, sheet_id, range_name, body)

    def create_and_import_data_to_google_sheet(self, data: list[list], options: dict):
        """
            Create new spreadsheet and import data to google drive
        Args:
            data (list[list]): Data to be imported
            options (dict): Options for creating spreadsheet and importing data
        Returns:
            googlesheet_url(str): Google Sheet URL

        """
        body_data = [list(data.keys())] + list(zip(*data.values()))
        format_cell_initial = {
            "backgroundColor": {
                "red": 0.0,
                "green": 0.0,
                "blue": 0.0
            },
            "horizontalAlignment": "CENTER",
            "textFormat": {
                "foregroundColor": {
                    "red": 1.0,
                    "green": 1.0,
                    "blue": 1.0
                },
                "fontSize": 12,
                "bold": True
            }
        }
        format_range_initial = {
            "sheetId": None,
            "startRowIndex": 0,
            "endRowIndex": 1,
            "startColumnIndex": 0,
            "endColumnIndex": len(body_data[0])
        }

        merge_range_initial = {
            "sheetId": None,
            "startRowIndex": 0,
            "endRowIndex": len(body_data),
            "startColumnIndex": 0,
            "endColumnIndex": 0,
        }

        spreadsheet = self.create_spreadsheet(options.get('spreadsheet_title'))
        spreadsheet_id = spreadsheet
        sheet_name = options.get('sheet_name')

        if sheet_name:
            sheet_id = self.create_sheet(spreadsheet_id, sheet_name)
        else:
            sheet_name = "Sheet1"
            sheet_id = self.get_sheet_id(spreadsheet_id, sheet_name="Sheet1")


        format_range = options.get('format_range') or format_range_initial
        merge_range = options.get('merge_range') or merge_range_initial
        format_cell = options.get('format_cell') or format_cell_initial
        format_range['sheetId'] = sheet_id
        merge_range['sheetId'] = sheet_id
        self.format_header(spreadsheet_id, format_cell, format_range)
        range_name = f"'{sheet_name}'!{chr(65)}1:{chr(65 + len(body_data[0]) - 1)}{len(body_data)}"

        self.merge_cell_and_write_data(spreadsheet_id, merge_range, range_name, body_data)

        googlesheet_url = f"{GOOGLE_SHEET_URL}/{spreadsheet_id}/edit#gid={sheet_id}"
        return googlesheet_url