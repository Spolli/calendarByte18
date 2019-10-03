import csv
from apiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
from datetime import datetime, timedelta

folder_path = {
    'csv': './src/res/cal.csv',
    'secret': './src/auth/client_secret.json'
}

def isint(x):
    try:
        a = float(x)
        b = int(a)
    except ValueError:
        return False
    else:
        return a == b


def getCredentials():
    scopes = ['https://www.googleapis.com/auth/calendar']
    flow = InstalledAppFlow.from_client_secrets_file(
        folder_path['secrete'], scopes=scopes)
    credentials = flow.run_console()
    pickle.dump(credentials, open("token.pkl", "wb"))
    credentials = pickle.load(open("token.pkl", "rb"))
    return build("calendar", "v3", credentials=credentials)

def formatDate(dt_str, hour):
    dt = dt_str[0].split('/')
    if not isint(hour):
        h = int(hour.split(',')[0])
        m = int(hour.split(',')[1])
        return datetime(int(dt[2]), int(dt[0]), int(dt[1]), h, m, 0, 0)
    return datetime(int(dt[2]), int(dt[0]), int(dt[1]), int(hour), 0, 0, 0)


def main():
    service = getCredentials()
    timezone = 'Europe/Rome'
    with open(folder_path['csv']) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        csv_reader.__next__()
        for row in csv_reader:
            event = {
                'summary': f'{row[4]}: {row[6]}',
                'location': row[8],
                'description': f'Numero modulo: {row[5]}',
                'start': {
                    'dateTime': formatDate(row, row[2].split('-')[0]),
                    'timeZone': timezone,
                },
                'end': {
                    'dateTime': formatDate(row, row[2].split('-')[1]),
                    'timeZone': timezone,
                },
                'reminders': {
                    'useDefault': True,
                    'overrides': [
                        {'method': 'popup', 'minutes': 10},
                    ],
                },
            }
            event = service.events().insert(calendarId='primary', body=event).execute()
            print(f'Event created: {event.get('htmlLink')}')

if __name__ == '__main__':
    main()
