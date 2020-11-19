import time
import os
import webbrowser
from typing import DefaultDict
from datetime import datetime
from programFiles.checkCredentials import checkCredentials
from programFiles.getResponses import getResponses
from programFiles.importantColFinder import getColLoc
from programFiles.excelManager import exportData
from programFiles.csv_processor import process_csv

# This file runs the overall program for downloading and exporting survey data.
# This program is only to be used for Question Pro responses uploaded to google sheets.


def checkNeccessaryFiles():
    '''This function checks to make sure that the neccessary files are there.
       If they are not it creates them and exits the program after telling the user to edit them.
    '''

    fileMissing = False

    if not os.path.exists('./programFiles/credentials.json'):
        fileMissing = True
        print('This program will now open an browser window to help setup.')
        print('Please go to the link and click enable the google sheets API.')
        print('Then download the configuration and place in into the programFiles folder for this application.')
        os.system('pause')
        webbrowser.open('https://developers.google.com/sheets/api/quickstart/python')

    if not os.path.exists('./resourceFiles/spreadsheet_ids.csv'):
        fileMissing = True

        with open('./resourceFiles/spreadsheet_ids.csv', 'w') as outputFile:
            print('Save Filepath (example user/downloads will create user/downloads/surveyName/),spreadsheet id,last-updated(null if never run)',
                  file=outputFile)

            print('Please edit the spreadsheet_ids.csv file and add the spreadsheet ids that you wish to grab responses from.')
            print('When finished re-run this program.\n')

    if not os.path.exists('./resourceFiles/answer_replacements.csv'):
        fileMissing = True

        with open('./resourceFiles/answer_replacements.csv', 'w') as outputFile:
            print('spreadsheet id,column name,"answer list","replacement list",County Question (1 if yes, 0 if no)',
                  file=outputFile)

            print('Please edit the answer_replacements.csv file and add the answers you want to be replaced.')
            print('When finished re-run this program.\n')

    if fileMissing:
        os.system('pause')
        exit(0)


def responseManager(creds):
    '''This function calls the functions to gets the answer replacements, and calls get responses.
       It then calls the export data to export the responses.
       And then replaces the last_updated for that response in the csv.
    '''

    csvData = process_csv('./resourceFiles/answer_replacements.csv')
    answerDict = {}
    countyQuestionDict = DefaultDict(lambda: '')

    for i in csvData:
        surveyId = i[0]

        # Skip title line and create answer replacement dictionary.
        if surveyId != 'spreadsheet id':
            colName = i[1]
            answerList = i[2].split(',')
            replacementList = i[3].split(',')

            if i[4] == '1':
                countyQuestionDict[surveyId] = colName

            if surveyId not in answerDict:
                answerDict[surveyId] = {}

            if colName not in answerDict[surveyId]:
                answerDict[surveyId][colName] = {
                    answerList[j]: replacementList[j] for j in range(len(answerList))}

    data = getResponses(creds)

    if not data:
        print('There was an error grabbing responses.  Make sure your spreadsheet ID list is not empty.')
        exit(1)

    for i in data:
        importantColumnDict = getColLoc(i[2][0], i[2][1])
        filepath = i[0][0]
        surveyId = i[0][1]
        surveyName = i[1]
        values = i[2][2:]
        lastUpdated = i[0][2]

        # Call export data with empty dictionary if this spreadsheet has no
        #   replacement answers.
        if i[0][1] in answerDict:
            exportData(filepath, surveyName, lastUpdated,
                       importantColumnDict, answerDict[surveyId], values, countyQuestionDict[surveyId])
        else:
            exportData(filepath, surveyName, lastUpdated,
                       importantColumnDict, {}, values, countyQuestionDict[surveyId])

        i[0][2] = datetime.now()

    if len(data) != 0:
        with open('./resourceFiles/spreadsheet_ids.csv', 'w') as outputFile:
            print('Save Filepath (example user/downloads will create user/downloads/surveyName/),spreadsheet id,last-updated(null if never run)',
                  file=outputFile)
            for i in data:
                print(','.join(str(j) for j in i[0]), file=outputFile)


if __name__ == "__main__":

    checkNeccessaryFiles()
    creds = checkCredentials()

    if not creds.valid:
        print('There was a problem with your credentials for Google sheets.')
        print('Please re-run this script and make sure to allow access to your Google sheets account.')
        os.system('pause')
        exit(1)

    while True:

        # Run response manager to get and handle responses. Then sleep 10 mins and do it again.
        responseManager(creds)
        time.sleep(600)
