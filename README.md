# QuestionProSurveyProject
Project to help a friend with re-organizing QuestionPro survey responses into a human readable format

#### In order for this program to be run the following steps must be done:
1. Install [python](https://www.python.org/downloads/).
2. Go [here](https://developers.google.com/sheets/api/quickstart/python) and get the credentials.json file from step 1 at the link and put it into the programFiles folder of this code.
3. Run the following command in a terminal window:
        pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
4. Then run the code in this project.  It should create a csv files for you to put survey information in, spreadsheet_ids.csv
    1. Fill this csv out with the details from the first row of the file.  Your spreadsheet id will be located in the url to your google sheet.
    2. Example: spreadsheets/d/**1UjPvBzWMZ1a8g45Tb8v96arKfb**/edit#gid=0
5. Run the program again and it should bring up a browser window to connect it to your google sheets account.
6. Now when you run the program one last time it will grab the sheets information from the sheed ids you provided and put it into a human readable format.
