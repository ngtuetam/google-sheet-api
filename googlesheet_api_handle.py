from googlesheet_api import GoogleSheetAPI


class HandleGoogleSheetAPI(GoogleSheetAPI):
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

    def create_and_import_data_to_google_sheet(self, data: list[list], format_cell: dict, format_range: dict, merge_range: dict) -> None:
        """
            Import data to google drive
        Args:
            data (list[list]): Data to be imported
            cell_format (dict): format cell
            format_cell (dict): dict format
        """

        spreadsheet = self.create_spreadsheet("New Spreadsheet")
        spreadsheet_id = spreadsheet

        worksheet_title = 'New Worksheet'
        sheet_id = self.create_sheet(spreadsheet_id, worksheet_title)

        format_range['sheetId'] = sheet_id
        merge_range['sheetId'] = sheet_id
        self.format_header(spreadsheet_id, format_cell, format_range)

        num_rows = len(data)
        num_columns = len(data[0])
        range_name = f"'{worksheet_title}'!{chr(65)}1:{chr(65 + num_columns - 1)}{num_rows}"

        self.merge_cell_and_write_data(spreadsheet_id, merge_range, range_name, data)