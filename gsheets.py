from __future__ import print_function
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


class Creds:
    def __init__(self):
        self.creds = None

    def OpenFile(self, filename):
        with open(filename, 'rb') as token:
            self.creds = pickle.load(token)

    def Valid(self):
        return self.creds or not self.creds.valid

    def needs_refresh(self):
        return self.creds and self.creds.expired and self.creds.refresh_token

    def refresh(self):
        self.creds.refresh(Request())

    def Login(self, credentials_file):
        flow = InstalledAppFlow.from_client_secrets_file(
                    credentials_file, SCOPES)
        self.creds = flow.run_local_server(port=0)

    def Save(self, filename):
        with open(filename, 'wb') as token:
            pickle.dump(self.creds, token)


class SpreadSheets():
    def __init__(self, creds):
        service = build('sheets', 'v4', credentials=creds)
        self.ss = service.spreadsheets()

    def read(self, sheet_id: str, range: str):
        return ValueRange(self.ss.values().get(spreadsheetId=sheet_id,
                          range=range).execute())

    def write(self, sheet_id, range, values):
        self.ss.values().update(spreadsheetId=sheet_id,
                                valueInputOption='USER_ENTERED',
                                range=range, body={'values': values}).execute()

    def batch_read(self, sheet_id, ranges):
        result = []
        vrs = self.ss.values().batchGet(spreadsheetId=sheet_id,
                                        ranges=ranges).execute()['valueRanges']
        for vr in vrs:
            result.append(ValueRange(vr))

        return result

    def batch_write(self, sheet_id, valueRanges):
        data = []
        for vr in valueRanges:
            data.append({'range': vr.range,
                         'values': vr.values})

        body = {'data': data, 'valueInputOption': 'USER_ENTERED'}

        self.ss.values().batchUpdate(spreadsheetId=sheet_id,
                                     body=body).execute()


class ValueRange:
    def __init__(self, value_range_json):
        self.range = value_range_json['range']
        self.values = value_range_json['values']

    def create(range, values):
        return ValueRange({'range': range, 'values': values})
