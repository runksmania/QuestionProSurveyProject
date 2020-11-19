import os
import sys
from googleapiclient.discovery import build
from programFiles.columnConversion import columnToLetter
from pathlib import Path


def getResponses(creds):
    '''Function to grab google sheets data from google api.
       Uses a csv file to get save filepath, sheet id, and last-updated.
       Then queries google api for data on survey google sheet.
       Returns a list that has the following:
            1. A list containing the information in the csv file for easy updating of csv.
            2. The name of the survey as give by the google sheets api.
            3. The data values.
    '''

    if creds.valid:

        if os.path.exists(os.path.abspath('./resourceFiles/spreadsheet_ids.csv')):
            data = []

            # Grab sheet ids from file.
            with open(os.path.abspath('./resourceFiles/spreadsheet_ids.csv'), 'r') as inputFile:
                lines = inputFile.readlines()

                for i in range(len(lines)):

                    # Check if line is not column names.
                    if i != 0:

                        # Split lines and strip \n character.
                        lineSplit = lines[i]

                        if lineSplit[-1] == '\n':
                            lineSplit = lineSplit[:-1]

                        lineSplit = lineSplit.split(',')

                        data.append(lineSplit)

            responses = []

            for i in data:
                filePath = i[0]
                sheetId = i[1]
                lastUpdated = i[2]

                # Begin API Code
                service = build('sheets', 'v4', credentials=creds)

                try:

                    res = service.spreadsheets().get(spreadsheetId=sheetId, fields='sheets(data/rowData/values/userEnteredValue,properties(index,sheetId,title))').execute()
                    sheetIndex = 0
                    sheetName = res['sheets'][sheetIndex]['properties']['title'].split('_')[0]
                    lastRow = len(res['sheets'][sheetIndex]['data'][0]['rowData'])
                    lastColumn = columnToLetter(max([len(e['values']) for e in res['sheets'][sheetIndex]['data'][0]['rowData'] if e]))

                    sheetRange = 'A1:' + lastColumn + str(lastRow)

                    sheet = service.spreadsheets()
                    result = sheet.values().get(spreadsheetId=sheetId, range=sheetRange).execute()
                    values = result.get('values', [])

                    responses.append([[filePath, sheetId, lastUpdated], sheetName, values])

                except:
                    e = sys.exc_info()[0]

                    with open('errors.txt', 'a') as errorFile:
                        print(e, file=errorFile)

                    print(
                        'There was an error attempting to get the responses for sheet id: ' + sheetId)
                    print('Please check to make sure the spreadsheet id is accurate.\n')

            return responses

    else:
        print('There was a problem with your credentials for Google sheets.')
        print('Please make sure to allow this script access to your Google sheets account.')
