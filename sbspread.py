from __future__ import print_function
import pickle
import os.path
import json
from googleapiclient import discovery
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from httplib2 import Http
from oauth2client import file, client, tools

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1pWeQRgt4gTaAbQGrRkQyEa7V5QtmBJ7oJHsczy4x4h4'
#SAMPLE_RANGE_NAME = '馬鹿１本部!A1:AA10'


def googleDrive():
    SCOPES = 'https://www.googleapis.com/auth/drive.readonly.metadata'
    store = file.Storage('storage.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = discovery.build('drive', 'v3', http=creds.authorize(Http()))

    #create service for google drive service api
    #service = build('drive', 'v3', credentials=creds)
    folder_id='1J8KH2FCN-m1FFhWs-LW2Qvvj-0wfx1XP'
    results = service.files().list(q="'1J8KH2FCN-m1FFhWs-LW2Qvvj-0wfx1XP' in parents").execute()
    items = results.get('files', [])

    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
            print(u'{0} ({1})'.format(item['name'], item['id']))

def googleSpreadsheet():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    #create service for google spreadsheet api
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    tempvalue=sheet.get(spreadsheetId=SAMPLE_SPREADSHEET_ID).execute()
    sheets=tempvalue.get('sheets',[])
    for title in sheets:
        #out=json.loads(title[0])
        titlename=title['properties']['title']
        if (str(titlename)).endswith('本部'):
            print(titlename)
            rangename=titlename
            result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=rangename).execute()
            values = result.get('values', [])

            if not values:
                print('No data found.')
            else:
                #print('Name:')
                for row in values:
                    # Print columns A and E, which correspond to indices 0 and 4.
                    if row[26]=='OK':
                        print('%s' % (row[2]))

if __name__ == '__main__':
    #main()
    googleDrive()
    #googleSpreadsheet()