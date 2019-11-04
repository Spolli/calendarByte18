#!/usr/bin/python
# -*- coding: utf-8 -*-

from __future__ import print_function
from datetime import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import csv
import sys
import pandas as pd

#Global Variables
#######################################################################
folder_path = {
    'csv': './src/res/cal.csv',
    'xls': sys.argv[1],
    'credentials': './src/auth/credentials.json',
    'token': './src/auth/token.pickle'
}
SCOPES = [
    'https://www.googleapis.com/auth/calendar', 
    'https://www.googleapis.com/auth/calendar.events'
]
calendar_name = 'byte18'
#######################################################################

def isint(x):
    try:
        a = float(x)
        b = int(a)
    except ValueError:
        return False
    else:
        return a == b

def xls2csv():
    if not os.path.exists(folder_path['csv']):
        xls_file = pd.read_excel(folder_path['xls'])
        xls_file.to_csv(folder_path['csv'], index=False)

def formatDate(dt_str, hour):
    dt = dt_str[0].split('-')
    if not isint(hour):
        h = int(hour.split(',')[0])
        m = int(hour.split(',')[1])
        return datetime(int(dt[0]), int(dt[1]), int(dt[2]), h, m, 0, 0)
    return datetime(int(dt[0]), int(dt[1]), int(dt[2]), int(hour), 0, 0, 0)

def main():
    creds = None
    if os.path.exists(folder_path['token']):
        with open(folder_path['token'], 'rb') as token:
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
        with open(folder_path['token'], 'wb') as token:
            pickle.dump(creds, token)
    service = build('calendar', 'v3', credentials=creds)
    print("Connected")
    ######################################################################
    print('Finding Calendar ID')
    calendarID = None
    page_token = None
    calendar_list = service.calendarList().list(
        pageToken=page_token).execute()
    for calendar_list_entry in calendar_list['items']:
        if calendar_list_entry['summary'] == calendar_name:
            calendarID = calendar_list_entry['id']
            print(f'Clearing Event from: {calendarID}')
            #service.calendars().clear(calendarId=calendarID).execute()
            break
    
    if not page_token:
        calendar = {"summary": calendar_name}
        created_calendar = service.calendars().insert(body=calendar).execute()
        calendarID = created_calendar['id']
    
    print(f'Calendar Id: {calendarID}')
    ######################################################################
    print('Uploading Events...')
    with open(folder_path['csv']) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        csv_reader.__next__()
        for row in csv_reader:
            event = {
                'id': 'byte18',
                'summary': f'{row[4]}: {row[6]}',
                'location': row[8],
                'description': f'Numero modulo: {row[5]}',
                'start': {
                    'dateTime':
                    formatDate(
                        row,
                        row[2].split('-')[0]).strftime("%Y-%m-%dT%H:%M:%S"),
                    'timeZone':
                    'Europe/Rome',
                },
                'end': {
                    'dateTime':
                    formatDate(
                        row,
                        row[2].split('-')[1]).strftime("%Y-%m-%dT%H:%M:%S"),
                    'timeZone':
                    'Europe/Rome',
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                        {
                            'method': 'popup',
                            'minutes': 90
                        },
                    ],
                },
            }
            created_event = service.events().insert(calendarId=calendarID, body=event).execute()
            print(created_event)


if __name__ == '__main__':
    xls2csv()
    main()
