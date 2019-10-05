#!/usr/bin/python
# -*- coding: utf-8 -*-

from __future__ import print_function
from datetime import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ['https://www.googleapis.com/auth/calendar']
folder_path = {
    'csv': './src/res/cal.csv',
    'xls': './src/res/Cal.OTTOBREallievi.xlsx',
    'credentials': './src/auth/credentials.json',
    'token': './src/auth/token.pickle'
}
service = None


def getCredentials():
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


def createCalendar(calendar_name):
    calendar = {"summary": calendar_name}
    created_calendar = service.calendars().insert(body=calendar).execute()
    return created_calendar['id']


def main():
    getCredentials()
    print("Connected")
    calendarId = createCalendar('test')
    print(f'Calendar ID = {calendaID}')


if __name__ == '__main__':
    main()
