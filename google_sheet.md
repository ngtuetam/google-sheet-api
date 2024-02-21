# Google Sheet API

## API
- Các hàm API


    | Function Name | Parameters | Description | Return |
    |----------|----------|----------|----------|
    | build_service   | api_name(str)<span style="color:red">\*</span> <br/> version_name(str)<span style="color:red">\*</span>| Build service và xác thực người dùng  | service | 
    | create_spreadsheet_and_import_table   | spreadsheet_title(str)<span style="color:red">\*</span><br/>  table_data(dict[list])<span style="color:red">\*</span><br/> sheet_name(str)<span>\</span><br/>| Tạo mới một `spreadsheet` và import table  | None | 
    | create_new_sheet_and_import_table | spreadsheet_id (str)<span style="color:red">\*</span> sheet_id (str)<span style="color:red">\*</span> sheet_name (str) <span style="color:red">\*</span> table_data(list[list])<span style="color:red">\*</span><br/> | Tạo mới một sheet và import table | None |
    | update_existing_sheet | spreadsheet_id (str)<span style="color:red">\*</span><br/>sheet_name (str)<span style="color:red">\*</span> table_data (dict[list])<span style="color:red">\*</span><br/> sheet_range (str)<span style="color:red">\*</span><br/> | Update một sheet | None |
    | write_sheet | spreadsheet_id (str)<span style="color:red">\*</span><br/>sheet_name (str)<span style="color:red">\*</span> table_data (dict[list])<span style="color:red">\*</span><br/> sheet_range (str)<span style="color:red">\*</span><br/> | Import data vào một sheet|None |



    ```python
    # Code Example
    handle_google_sheet_api = HandleGoogleSheetAPI()

    # Dữ liệu
    table_data = {
        "Column1": [7, 8, 9],
        "Column2": ["P", "Q", "R"],

    }
    spreadsheet_title = "New Spreadsheet"

    # Thực hiện import dữ liệu đến google sheet
    handle_google_sheet_api.create_spreadsheet_and_import_table(spreadsheet_title, table_data)

    ```
