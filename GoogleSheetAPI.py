class GoogleSheetAPI():
    def __init__(self, json_cred_path):
        self.json_cred_path = json_cred_path
        self.creds = None
        self.service = None
        self.sheet = None
        self.SAMPLE_SPREADSHEET_ID = None
        self.SAMPLE_RANGE_NAME = None
        self.body = None
        self.request = None
        self.result = None
        self.values = None

    def init_service(self):
        self