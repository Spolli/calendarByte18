from __future__ import print_function
from datetime import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import csv
import xlrd

folder_path = {
    'csv': './src/res/cal.csv',
    'xls': './src/res/*.xlsx',
    'credentials': './src/auth/credentials.json',
    'token': './src/auth/token.pickle'
}

def csv_from_excel():
    #if not os.path.exists(folder_path['csv']):
    wb = xlrd.open_workbook(folder_path['xls'])
    sh = wb.sheet_by_name('Sheet1')
    your_csv_file = open(folder_path['csv'], 'wb')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
    for rownum in xrange(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()


def formatDate(dt_str, hour):
    dt = dt_str[0].split('/')
    if not isint(hour):
        h = int(hour.split(',')[0])
        m = int(hour.split(',')[1])
        return datetime(int(dt[2]), int(dt[0]), int(dt[1]), h, m, 0, 0)
    return datetime(int(dt[2]), int(dt[0]), int(dt[1]), int(hour), 0, 0, 0)


def isint(x):
    try:
        a = float(x)
        b = int(a)
    except ValueError:
        return False
    else:
        return a == b


def main():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                folder_path['credentials'], SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('calendar', 'v3', credentials=creds)
    print("Connected")
    with open(folder_path['csv']) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        csv_reader.__next__()
        for row in csv_reader:
            event = {
                'summary': f'{row[4]}: {row[6]}',
                'location': row[8],
                'description': f'Numero modulo: {row[5]}',
                'start': {
                    'dateTime': formatDate(row, row[2].split('-')[0]).strftime("%Y-%m-%dT%H:%M:%S"),
                    'timeZone': 'Europe/Rome',
                },
                'end': {
                    'dateTime': formatDate(row, row[2].split('-')[1]).strftime("%Y-%m-%dT%H:%M:%S"),
                    'timeZone': 'Europe/Rome',
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                        {'method': 'popup', 'minutes': 90},
                    ],
                },
                'recurrence': [
                ],
                'attendees': [
                ],
            }
            created_calendar = service.events().insert(
                calendarId='primary', body=event).execute()


if __name__ == '__main__':
    main()
