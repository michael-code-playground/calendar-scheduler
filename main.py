from google.oauth2 import service_account
from googleapiclient.discovery import build
import datetime
from googleapiclient.errors import HttpError
from dateutil.relativedelta import relativedelta
import os
def initiate_calendar():
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    SERVICE_ACCOUNT_FILE = 'the-method-427521-p2-74db9b4d1aa9.json'

    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES,
    )

    service = build('calendar', 'v3', credentials=credentials)
    return service

def access_sheet():
    
    SERVICE_ACCOUNT_FILE = 'the-method-427521-p2-74db9b4d1aa9.json'
    
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    return creds

def read_sheet(creds, spreadsheet_id, range_name):
    """Reads data from a Google Sheet."""
    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get("values", [])
        
        if not values:
         print("No data found.")

        return values
    except HttpError as err:
        print(f"An error occurred: {err}")
        return None

def insert_event(record,datetime):
    event = {
    'summary': record,
    'start': {
        'date': datetime,
        'timeZone': 'Europe/Lisbon',
    },
    'end': {
    'date': datetime,
    'timeZone': 'Europe/Lisbon',
    },
    'reminders': {
        'useDefault': False,
        'overrides': [
        {'method': 'email', 'minutes': 24 * 60},
        {'method': 'popup', 'minutes': 10},
        ],
    },
    }

    event = service.events().insert(calendarId=os.getenv("calendarid"), body=event).execute()


def format_date(event_date):
    date = datetime.datetime.strptime(event_date, "%d/%m/%Y").date().isoformat()
    return date

current_date = datetime.datetime.now(tz=datetime.timezone.utc)
two_months_later = current_date + relativedelta(months=2)


start_time = datetime.datetime.now(tz=datetime.timezone.utc).isoformat()
end_time = two_months_later.isoformat()

service = initiate_calendar()

events_result = (
    service.events()
    .list(
        calendarId=os.getenv("calendarid"),
        timeMin= start_time,
        timeMax= end_time,
        singleEvents=True,
        orderBy="startTime",
        )
        .execute()
    )
events = events_result.get("items", [])

if not events:
    print("No upcoming events found.")


calendar_event_keys = {(event["start"].get("dateTime", event["start"].get("date"))[:10], event['summary']) for event in events}
print("Calendar - current events:")
print()
print(calendar_event_keys)

creds = access_sheet()

SAMPLE_SPREADSHEET_ID = "1V9mLaKA9WhsyA_qcEf0Uy6KB2fGLkKI8IzhtiVTVm8Y"
SAMPLE_RANGE_NAME = "Activities!A2:B100"
values = read_sheet(creds, SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME)

print("Events in sheet:")
print()
sheet_event_keys = {(format_date(value[1]), value[0]) for value in values}
print(sheet_event_keys)
print("Events to create:")
print()
new_events_to_create = sheet_event_keys - calendar_event_keys
print(new_events_to_create, len(new_events_to_create))

for event in new_events_to_create:
    insert_event(event[1],event[0])
    