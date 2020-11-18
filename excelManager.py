import os
import csv
from openpyxl import Workbook
from datetime import datetime

# Header column dict containing lists of the following for excel printing:
#   1. Where to print the column name.
#   2. Where to print their value.
headerColumns = {
    'First Name': ['A1', 'B1'],
    'Last Name': ['C1', 'D1'],
    'Phone': ['A2', 'B2'],
    'Email Address': ['C2', 'D2'],
    'Address Line 1': ['A3', 'B3'],
    'Address Line 2': ['A4', 'B4'],
    'City': ['C3', 'D3'],
    'State': ['C4', 'D4'],
    'Zipcode': ['E4', 'F4'],
    'Response ID': ['A6', 'B6'],
    'IP Address': ['A7', 'B7'],
    'Device': ['A8', 'B8'],
    'Operating System': ['A9', 'B9'],
    'Language': ['A10', 'B10'],
    'Timestamp (mm/dd/yyyy)': ['C6', 'D6'],
    'Country Code': ['C7', 'D7'],
    'Region': ['C8', 'D8'],
    'Browser': ['C9', 'D9'],
    'Time Taken to Complete (Seconds)': ['E6', 'F6'],
    'Respondent Email': ['E8', 'F8'],
}


def getAnswer(answer, colName, answerDict):

    # Replacment answer function.
    if colName in answerDict:

        if answer in answerDict[colName]:

            # If answer and colName found return its replacement answer.
            return answerDict[colName][answer]

    # Otherwise return original answer
    return answer


def exportData(filepath, surveyName, lastUpdated, columnDict, answerDict, data, countyQuestion):
    '''Function to manage exporting the data to the excel file.'''

    # Create filepaths if they do not exist.
    if not os.path.exists(filepath):
        os.mkdir(filepath)

    filepath += '/' + surveyName + '/'

    if not os.path.exists(filepath):
        os.mkdir(filepath)

    if os.path.exists(filepath):

        # If previously updated parse lastUpdated datetime string.
        if lastUpdated != 'null':
            lastUpdated = datetime.strptime(
                lastUpdated, "%Y-%m-%d %H:%M:%S.%f")

        for i in data:

            # Parse datetime string of completion of the survey response.
            dataTimestamp = datetime.strptime(
                i[columnDict['Timestamp (mm/dd/yyyy)'][0]], '%a %d %b, %H:%M:%S %Z %Y')

            # If survey has never been updated or updated prior to survey response then export data.
            if lastUpdated == 'null' or lastUpdated < dataTimestamp:
                wb = Workbook()
                sheet = wb.active

                # Create saveName variable for saving excel sheet later.
                # Starts with response id for unique naming scheme.
                # Then adds county and last name of survey taker.
                saveName = i[columnDict['Response ID'][0]]

                if countyQuestion in columnDict and countyQuestion != '':
                    saveName += '-' + getAnswer(i[columnDict[countyQuestion][0]],
                                                countyQuestion, answerDict)

                if 'Last Name' in columnDict and (i[columnDict['Last Name'][0]] != '' or i[columnDict['Last Name'][0]] != ' '):
                    saveName += '-' + i[columnDict['Last Name'][0]]

                row = 15
                sheet['A12'] = 'Category:'

                for colName, v in columnDict.items():

                    if colName in headerColumns:

                        # If column is a header column print it and its answer in the appropriate spot.
                        sheet[headerColumns[colName][0]] = colName + ':'
                        sheet[headerColumns[colName][1]] = getAnswer(
                            i[v[0]], colName, answerDict)

                    else:

                        # Not a header column and is a question column.
                        # Get sub column list and instantiate boolean checks.
                        subColList = columnDict[colName][1]
                        dataAdded = False
                        questionPrinted = False

                        if len(subColList) > 0:

                            # If there is subcolumns iterate through the list and print answers
                            for j in range(len(subColList)):
                                subColumnName = subColList[j]

                                if not v[0] + j >= len(i):

                                    # Check to make sure answers are not blank and if subcolumn is other.
                                    # Other is printed differently by questionPro than other subcolumns.
                                    if subColumnName == 'Other' and i[v[0] + j + 1] != '' and i[v[0] + j + 1] != ' ' and len(subColList) == 1:

                                        if i[v[0] + j + 1] != '' or i[v[0] + j + 1] != ' ':
                                            dataAdded = True

                                            if not questionPrinted:
                                                sheet['A' +
                                                      str(row)] = 'Question: '
                                                sheet['B' + str(row)] = colName
                                                row += 1
                                                questionPrinted = True

                                            sheet['A' +
                                                  str(row)] = subColumnName
                                            sheet['B' +
                                                  str(row)] = getAnswer(i[v[0] + j + 1], subColumnName, answerDict)
                                            row += 1

                                    elif subColumnName == 'Other' and len(subColList) == 1:

                                        if i[v[0]] != '' or i[v[0]] != ' ':
                                            dataAdded = True

                                            if not questionPrinted:
                                                sheet['A' +
                                                      str(row)] = 'Question: '
                                                sheet['B' + str(row)] = colName
                                                row += 1
                                                questionPrinted = True

                                            sheet['A' + str(row)] = 'Answer: '
                                            sheet['B' + str(row)] = getAnswer(i[v[0]], subColumnName, answerDict)
                                            row += 1

                                    elif i[v[0] + j] != '' and i[v[0] + j] != ' ':

                                        if i[v[0] + j] != '' or i[v[0] + j] != ' ':
                                            dataAdded = True

                                            if not questionPrinted:
                                                sheet['A' +
                                                      str(row)] = 'Question: '
                                                sheet['B' + str(row)] = colName
                                                row += 1
                                                questionPrinted = True

                                            sheet['A' +
                                                  str(row)] = subColumnName
                                            sheet['B' + str(row)] = getAnswer(i[v[0] + j], subColumnName, answerDict)
                                            row += 1

                        else:

                            # If no subcolumn list check for non blanks and print answers.
                            if v[0] < len(i) and i[v[0]] != '' and i[v[0]] != ' ':
                                dataAdded = True

                                if not questionPrinted:
                                    sheet['A' +
                                          str(row)] = 'Question: '
                                    sheet['B' + str(row)] = colName
                                    row += 1
                                    questionPrinted = True

                                sheet['A' + str(row)] = 'Answer: '
                                sheet['B' + str(row)] = getAnswer(i[v[0]], colName, answerDict)
                                row += 1

                        if dataAdded:
                            row += 4

                wb.save(filepath + saveName + '.xlsx')
