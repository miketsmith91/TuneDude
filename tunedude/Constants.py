import json

with open('config.json') as json_data_file:
    config = json.load(json_data_file)

SCOPES = config['DEFAULT']['SCOPES']
CLIENT_SECRET_FILE = config['DEFAULT']['CLIENT_SECRET_FILE']
APPLICATION_NAME = config['DEFAULT']['APPLICATION_NAME']
SPREADSHEET_ID = config['DEFAULT']['SPREADSHEET_ID']
ABSOLUTE_PATH = config['DEFAULT']['ABSOLUTE_PATH']
ITUNES_PATH = config['DEFAULT']['ITUNES_PATH']

TEST_SPREADSHEET_ID = config['TEST']['TEST_SPREADSHEET_ID']