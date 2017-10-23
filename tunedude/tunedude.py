from __future__ import print_function
from __future__ import unicode_literals
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

# Imports Below are for the testing to see if the rest of the program works with Google API
import youtube_dl
import sys
import subprocess as sp
# Imports below are for moving converted file from Program Directory to iTunes Directory
import fnmatch
import shutil
import eyed3
import datetime
import json

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json

with open('config.json') as json_data_file:
    config = json.load(json_data_file)

SCOPES = config['DEFAULT']['SCOPES']
CLIENT_SECRET_FILE = config['DEFAULT']['CLIENT_SECRET_FILE']
APPLICATION_NAME = config['DEFAULT']['APPLICATION_NAME']
SPREADSHEET_ID = config['DEFAULT']['SPREADSHEET_ID']
ABSOLUTE_PATH = config['DEFAULT']['ABSOLUTE_PATH']

def get_credentials():
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
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials






def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    os.chdir(ABSOLUTE_PATH)
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    rangeName = 'In Progress!A2:E'
    def getInProgressSpreadSheet():
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=rangeName).execute()
        values = result.get('values', [])
        return values

    def addToErrorSpreadsheet(row,rowNumber,errorName,errorDesc):
        while len(row) < 3:
            row.append("NOT ENTERED")
        row.append(errorName)
        row.append(str(errorDesc))
        value_range_body = {
            # TODO: Add desired entries to the request body.
            "values": [row]
        }
        appendToErrorRequest = service.spreadsheets().values().append(spreadsheetId=SPREADSHEET_ID, range=u"Error!A2:E",
                                                                      valueInputOption='USER_ENTERED',
                                                                      body=value_range_body)
        response = appendToErrorRequest.execute()

        clearRange = u"In Progress!A%s:C%s" % (rowNumber, rowNumber)

        # clear_body = {
        #     "range": clearRange,
        #     "values": [[]]
        #
        # }
        requests = []
        # requests.append({
        #     'deleteRange': {
        #         "range": clearRange,
        #         "shiftDimension": "ROWS",
        #     }
        # })
        requests.append({
            'deleteDimension': {
                "range": {
                    "sheetId": 0,
                    "dimension": "ROWS",
                    "startIndex": rowNumber-1,
                    "endIndex": rowNumber,}
            }
        })
        body = {
            'requests': requests
        }

        deleteRangeRequest = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID,
                                                                    body=body)

        deleteResponse = deleteRangeRequest.execute()

        dumbVariable = "nothing"

    values = getInProgressSpreadSheet()

    if not values:
        print('No data found.')
    else:
        print('Link, Song Artist, Song Name:')
        rowNumber = 2
        for row in values:
            try:

                # Print columns A and E, which correspond to indices 0 and 4.
                print('%s, %s, %s' % (row[0], row[1], row[2]))
                songLink = row[0]
                songArtist = row[1]
                songName = row[2]
                youtube_dl_Output_Name = songArtist + " - " + songName
                completeSheetID = 1911851509
            except Exception as err:
                print ("An exception just occured dude!: " +str(err))
                exception = "Empty Non-optional Values"
                addToErrorSpreadsheet(row,rowNumber,exception,err)
                values = getInProgressSpreadSheet()
                continue
            options = {
                # Changed to only bestaudio because i'm skeptical. 'format': 'bestaudio/best', # choice of quality
                'format': 'bestaudio',
                'extractaudio': True,  # only keep the audio. No video needed.
                'audioformat': "mp3",  # convert to mp3
                'outtmpl': youtube_dl_Output_Name+'.%(ext)s',  # name the file the ID of the video
                'noplaylist': True,}  # only download single song, not playlist

            try:

                ydl = youtube_dl.YoutubeDL(options)

                result = ydl.extract_info(songLink, download=True)

            except Exception as err:
                print ("An exception just occurred bro!: " + str(err))
                # row[4] = err
                exception = 'Invalid Link'
                addToErrorSpreadsheet(row,rowNumber,exception,err)
                values = getInProgressSpreadSheet()

                continue

            print(result.get('title'))

            songFileName = youtube_dl_Output_Name+"."+result.get('ext')
            print(songFileName)

            if sys.platform == "Windows":
                FFMPEG_BIN = "ffmpeg.exe"
            else:
                FFMPEG_BIN = "/usr/local/bin/ffmpeg"

            songFileNameOutput = row[1]+' - '+row[2]+".mp3"

            command = [ FFMPEG_BIN,
                        '-i', songFileName,
                        '-b:a', '320K',
                        #'-vn', row[1]+' - '+row[2]+".mp3"]
                        '-vn', songFileNameOutput]
            pipe = sp.Popen(command, stdout=sp.PIPE)

            pipe.wait()

            pipe.stdout.close()

            dst = "/Users/mike/Music/iTunes/iTunes Media/Automatically Add to iTunes"

            for file in os.listdir(ABSOLUTE_PATH):
                if fnmatch.fnmatch(file, '*.mp3'):
                    print (file)

                    src = ABSOLUTE_PATH

                    # Now modify ID3 metadata of file provided in Google Sheet
                    audioFile = eyed3.load(file)
                    audioFile.tag.artist = songArtist
                    audioFile.tag.title = songName
                    if len(row) > 3:

                        if len(row) == 4:
                            albumArtist = row[3]
                            audioFile.tag.album_artist = albumArtist
                        elif row[3] == u'':
                            albumName = row[4]
                            audioFile.tag.album = albumName
                        else:
                            albumArtist = row[3]
                            albumName = row[4]
                            audioFile.tag.album_artist = albumArtist
                            audioFile.tag.album = albumName

                    audioFile.tag.save()

                    path = os.path.join(src, file)
                    shutil.move(path, dst)

                    #response = request.execute()

                if fnmatch.fnmatch(file, "*.webm"):
                    print(file+" has been deleted.")
                    os.remove(file)
            rowNumber += 1
        st = str(datetime.datetime.now()).split('.')[0]
        for row in values:
            while len(row) < 5:
                row.append('*LEFT BLANK*')
            row.append(st)

        #Update Completed Sheet

        # The A1 notation of a range to search for a logical table of data.
        # Values will be appended after the last row of the table.
        range_ = u"Completed!A2:E"

        # How the input data should be interpreted.
        value_input_option = 'USER_ENTERED'

        # How the input data should be inserted.
        insert_data_option = u'INSERT_ROWS'

        value_range_body = {
            "values": values
        }

        appendRequest = service.spreadsheets().values().append(spreadsheetId=SPREADSHEET_ID, range=range_,
                                                                           valueInputOption=value_input_option,
                                                                           body=value_range_body)
        response = appendRequest.execute()

        clearRequest = service.spreadsheets().values().clear(spreadsheetId=SPREADSHEET_ID,
                                                                        range=rangeName,
                                                                        body={})
        clearResponse = clearRequest.execute()

        # LOGGER UNDER CONSTRUCTION ####

        # import logging
        # logging.basicConfig(filename='myProgramLog.txt', level=logging.DEBUG, format='
        #                                                                              % (asctime)
        # s - %(levelname)
        # s - % (message)
        # s
        # ')
# TODO: Create Unit Tests
# TODO: Create Logger

if __name__ == '__main__':
    main()