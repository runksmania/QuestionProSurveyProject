import csv


def process_csv(fileName):
    data = []

    with open(fileName, 'r') as someFile:
        csvFile = csv.reader(someFile, delimiter=',', quotechar='"')

        for row in csvFile:
            data.append(row)

    return data
