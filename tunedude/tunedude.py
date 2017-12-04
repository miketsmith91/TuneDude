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
from Constants import SCOPES,CLIENT_SECRET_FILE,APPLICATION_NAME,SPREADSHEET_ID,ABSOLUTE_PATH,ITUNES_PATH, TEST_SPREADSHEET_ID
import SheetService

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json

def main():
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """

    os.chdir(ABSOLUTE_PATH)
    inProgressRange = 'In Progress!A2:E'

    GoogleSheet = SheetService.SheetService(TEST_SPREADSHEET_ID)
    values = GoogleSheet.getSpreadsheet(inProgressRange)

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
                #addToErrorSpreadsheet(row,rowNumber,exception,err)
                GoogleSheet.addToErrorSpreadsheet(row,rowNumber,exception,err)
                #values = getInProgressSpreadSheet()
                values = GoogleSheet.getSpreadsheet(inProgressRange)
                continue
            options = {
                # Changed to only bestaudio because i'm skeptical. 'format': 'bestaudio/best', # choice of quality
                'format': 'bestaudio',
                'extractaudio': True,  # only keep the audio. No video needed.
                'audioformat': "mp3",  # convert to mp3
                'outtmpl': youtube_dl_Output_Name+'.%(ext)s',  # name the file the ID of the video
                #'outtmpl': '%(title)s' + '.%(ext)s',  # name the file the ID of the video
                'noplaylist': True,}  # only download single song, not playlist

            try:

                ydl = youtube_dl.YoutubeDL(options)

                result = ydl.extract_info(songLink, download=True)

            except Exception as err:
                print ("An exception just occurred bro!: " + str(err))
                # row[4] = err
                exception = 'Invalid Link'
                #addToErrorSpreadsheet(row,rowNumber,exception,err)
                GoogleSheet.addToErrorSpreadsheet(row,rowNumber,exception,err)
                #values = getInProgressSpreadSheet()
                values = GoogleSheet.getSpreadsheet(inProgressRange)

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

            dst = ITUNES_PATH

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
                    if os.path.exists(dst):
                        shutil.move(path, dst)

                    #response = request.execute()

                if fnmatch.fnmatch(file, "*.webm"):
                    print(file+" has been deleted.")
                    os.remove(file)
            rowNumber += 1
        st = str(datetime.datetime.now()).split('.')[0]

        GoogleSheet.updateCompleteSheet(values,st)
        GoogleSheet.clearSheet('In Progress!A2:E')

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