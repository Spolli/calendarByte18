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
    'xls': './src/res/Cal.OTTOBREallievi.xlsx',
    'credentials': './src/auth/credentials.json',
    'token': './src/auth/token.pickle'
}

calendaID = 'qrm4o759mdernjd4j8f295sh98@group.calendar.google.com'

try:
    folder_path['xls'] = sys.argv[1]
except Exception as e:
    folder_path['xls'] = './src/res/Cal.OTTOBREallievi.xlsx'

SCOPES = ['https://www.googleapis.com/auth/calendar']
service = None
#######################################################################


def xls2csv():
    xls_file = pd.read_excel(folder_path['xls'])
    xls_file.to_csv(folder_path['csv'], index=False)


def formatDate(dt_str, hour):
    dt = dt_str[0].split('-')
    if not isint(hour):
        h = int(hour.split(',')[0])
        m = int(hour.split(',')[1])
        return datetime(int(dt[0]), int(dt[1]), int(dt[2]), h, m, 0, 0)
    return datetime(int(dt[0]), int(dt[1]), int(dt[2]), int(hour), 0, 0, 0)


def isint(x):
    try:
        a = float(x)
        b = int(a)
    except ValueError:
        return False
    else:
        return a == b


def existCalendar(calendar_name):
    page_token = None
    while True:
        calendar_list = service.calendarList().list(
            pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            if calendar_list_entry['summary'] == calendar_name:
                return calendar_list_entry['id']
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            return False


def createCalendar(calendar_name):
    calendar = {"summary": calendar_name}
    created_calendar = service.calendars().insert(body=calendar).execute()
    return created_calendar['id']


def main():
    global service
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
    calendar = {"summary": "byte18"}
    created_calendar = service.calendars().insert(body=calendar).execute()
    print()
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
                'recurrence': [],
                'attendees': [],
            }
            created_calendar = service.events().insert(
                calendarId=created_calendar['id'], body=event).execute()


if __name__ == '__main__':
    xls2csv()
    main()
