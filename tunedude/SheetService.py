import os
import httplib2
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

from Constants import SCOPES,CLIENT_SECRET_FILE,APPLICATION_NAME
import datetime


class SheetService(object):
    def __init__(self, spreadsheet_id):
        self.spreadsheet_id = spreadsheet_id
        credentials = self.__get_credentials()
        http = credentials.authorize(httplib2.Http())
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                        'version=v4')
        service = discovery.build('sheets', 'v4', http=http,
                                  discoveryServiceUrl=discoveryUrl)
        self.service = service

    def __get_credentials(self):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir,
                                       'sheets.googleapis.com-python-quickstart.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
            flow.user_agent = APPLICATION_NAME
            if flags:
                credentials = tools.run_flow(flow, store, flags)
            else:  # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def getSpreadsheet(self, my_range):
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id, range=my_range).execute()
        values = result.get('values', [])
        return values

    def delete_entry(self, rowNumber):
        requests = []
        requests.append({
            'deleteDimension': {
                "range": {
                    "sheetId": 0,
                    "dimension": "ROWS",
                    "startIndex": rowNumber - 1,
                    "endIndex": rowNumber, }
            }
        })
        body = {
            'requests': requests
        }
        deleteRangeRequest = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheet_id,
                                                                body=body)
        deleteRangeRequest.execute()

    def addToSpreadsheet(self, range, row):
        value_range_body = {
            "values": [row]
        }
        appendToRequest = self.service.spreadsheets().values().append(spreadsheetId=self.spreadsheet_id, range=range,
                                                                 valueInputOption='USER_ENTERED',
                                                                 body=value_range_body)
        appendToRequest.execute()

    def addToErrorSpreadsheet(self, row, rowNumber,errorName, errorDesc):
        while len(row) < 3:
            row.append("NOT ENTERED")
        row.append(errorName)
        row.append(str(errorDesc))
        self.addToSpreadsheet(u"Error!A2:E",row)
        self.delete_entry(rowNumber)

    def updateCompleteSheet(self,values,timestamp):
        #st = str(datetime.datetime.now()).split('.')[0]
        for row in values:
            while len(row) < 5:
                row.append('*LEFT BLANK*')
            row.append(timestamp)

        # Update Completed Sheet

        # The A1 notation of a range to search for a logical table of data.
        # Values will be appended after the last row of the table.
        range_ = u"Completed!A2:E"

        # How the input data should be interpreted.
        value_input_option = 'USER_ENTERED'

        # How the input data should be inserted.
        #insert_data_option = u'INSERT_ROWS'

        value_range_body = {
            "values": values
        }

        appendRequest = self.service.spreadsheets().values().append(spreadsheetId=self.spreadsheet_id, range=range_,
                                                               valueInputOption=value_input_option,
                                                               body=value_range_body)
        appendRequest.execute()

    def clearSheet(self,range):

        clearRequest = self.service.spreadsheets().values().clear(spreadsheetId=self.spreadsheet_id,
                                                             range=range,
                                                             body={})
        clearRequest.execute()