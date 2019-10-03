import csv
from apiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
from datetime import datetime, timedelta

folder_path = {
    'csv': './res/cal.csv',
    'secret': './src/auth/client_secret.json'
}


def getCredentials():
    scopes = ['https://www.googleapis.com/auth/calendar']
    flow = InstalledAppFlow.from_client_secrets_file(
        folder_path['secrete'], scopes=scopes)
    credentials = flow.run_console()
    pickle.dump(credentials, open("token.pkl", "wb"))
    credentials = pickle.load(open("token.pkl", "rb"))
    return build("calendar", "v3", credentials=credentials)


def main():
    service = getCredentials()
    result = service.calendarList().list().execute()
    with open(folder_path['csv']) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        for row in csv_reader:
            event = {
                'summary': f'{row[4]}: {row[6]}',
                'location': row[8],
                'description': f'Numero modulo: {row[5]}',
                'start': {
                    'dateTime': start_time.strftime("%Y-%m-%dT%H:%M:%S"),
                    'timeZone': timezone,
                },
                'end': {
                    'dateTime': end_time.strftime("%Y-%m-%dT%H:%M:%S"),
                    'timeZone': timezone,
                },
                'reminders': {
                    'useDefault': False,
                    'overrides': [
                        {'method': 'email', 'minutes': 24 * 60},
                        {'method': 'popup', 'minutes': 10},
                    ],
                },
            }


if __name__ == '__main__':
    main()
