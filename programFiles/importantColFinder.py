from collections import OrderedDict


def getColLoc(columnList, subColumnList):
    '''Function to look through the column & subcolumn list and extract the important ones.
        Returns a dictionary of the following:
            A key that is the master column
            A value that is of list type which contains the following:
                1. The location of the master column in the column list to easily grab important data.
                2. A list containing the subcolumns for the master column.
    '''

    # Columns we don't care about.
    colSkipDict = {
        'Duplicate': 0,
        'Seq. Number': 0,
        'External Reference': 0,
        'Email List': 0,
        'Response Status': 0,
        'Country': 0,
        'Longitude': 0,
        'Latitude': 0,
        'Radius': 0,
        '': 0
    }

    colKeepDict = OrderedDict()
    i = 0

    while i < (len(columnList)):
        colName = columnList[i]

        if i == 337:
            print()

        # Check if its a column we don't want to skip and push if it is.
        if colName not in colSkipDict:
            colNameSplit = columnList[i].split()

            if colName != '' and colNameSplit[0] != 'Custom':
                colKeepDict[colName] = [i, []]

        # If there are sub columns for this item push them into array for the key.
        if i < len(subColumnList) and subColumnList[i] != '':
            masterColName = columnList[i]

            if masterColName in colKeepDict:
                j = i + 1
                colKeepDict[masterColName][1].append(subColumnList[i])

                while j < len(subColumnList) and subColumnList[j] != '':

                    if j >= len(columnList) or columnList[j] == '':
                        colKeepDict[masterColName][1].append(subColumnList[j])

                    else:
                        i = j - 1
                        break

                    j += 1

                    i = j - 1

            # The other sub column is positioned differently.  Check for it and push into array for key.
            elif i < len(subColumnList) and (subColumnList[i] == 'Other' or subColumnList[i] == 'Not Listed' or subColumnList[i] == 'Dynamic Comment'):
                back_count = i - 1

                while columnList[back_count] == '':
                    back_count -= 1

                if columnList[back_count] in colKeepDict:
                    colKeepDict[columnList[back_count]][1].append(subColumnList[i])

        i += 1

    return colKeepDict
