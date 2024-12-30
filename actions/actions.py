from __future__ import print_function

from typing import Any, Text, Dict, List
from rasa_sdk import Action, Tracker
from rasa_sdk.executor import CollectingDispatcher
from rasa_sdk.events import AllSlotsReset

import datetime
from datetime import datetime, timedelta
import os.path
import pickle
import pytz  # Ensure pytz is installed: pip install pytz

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ['https://www.googleapis.com/auth/calendar']
CREDENTIALS_FILE = 'credentials.json'

def get_calendar_service():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
            print("Loaded credentials from token.pickle")

    # If no valid credentials, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            print("Credentials refreshed")
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
            print("Obtained new credentials")

        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
            print("Saved new credentials to token.pickle")

    try:
        service = build('calendar', 'v3', credentials=creds)
        return service
    except Exception as e:
        print(f"An error occurred while building the service: {e}")
        return None

def add_event(service, event_name: Text, event_start: datetime, event_end: datetime):
    """
    Adds a new event to Google Calendar.
    """
    try:
        new_event = {
            'summary': event_name,
            'location': "Default Location",  # Customize as needed
            'description': "Automatically added event",
            'start': {
                'dateTime': event_start.isoformat(),
                'timeZone': 'Asia/Ho_Chi_Minh',
            },
            'end': {
                'dateTime': event_end.isoformat(),
                'timeZone': 'Asia/Ho_Chi_Minh',
            },
            'reminders': {
                'useDefault': True,
            },
        }
        created_event = service.events().insert(calendarId='primary', body=new_event).execute()
        print(f"Created event: {created_event.get('htmlLink')}")
        return created_event
    except HttpError as error:
        print(f"An error occurred while adding the event: {error}")
        return None

def get_events(service, time_min: Text, time_max: Text) -> List[Dict[str, Any]]:
    """
    Retrieves events within the specified time range.
    """
    try:
        events_result = service.events().list(
            calendarId='primary',
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        events = events_result.get('items', [])
        return events
    except HttpError as error:
        print(f"An error occurred while fetching events: {error}")
        return []

class AddEventToCalendar(Action):

    def name(self) -> Text:
        return "action_add_event"

    def run(self, dispatcher: CollectingDispatcher,
            tracker: Tracker,
            domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        event_name = tracker.get_slot('event')
        time_str = tracker.get_slot('time')

        if not event_name or not time_str:
            dispatcher.utter_message(text="Please provide both the event name and time.")
            return [AllSlotsReset()]

        try:
            # Parse the time slot to a datetime object
            event_start_time = datetime.strptime(time_str, '%d/%m/%y %H:%M:%S')

            # Assign timezone to the datetime object
            tz = pytz.timezone('Asia/Ho_Chi_Minh')
            event_start_time = tz.localize(event_start_time)

            # Calculate event end time (1 hour duration)
            event_end_time = event_start_time + timedelta(hours=1)

            # Define time range for conflict checking (same as event time)
            time_min = event_start_time.isoformat()
            time_max = event_end_time.isoformat()

            # Initialize Google Calendar API
            service = get_calendar_service()
            if not service:
                dispatcher.utter_message(text="Failed to initialize Google Calendar service.")
                return [AllSlotsReset()]

            # Retrieve existing events in the desired time slot
            existing_events = get_events(service, time_min, time_max)

            # Initialize conflict flag
            conflict = False
            conflicting_event = None

            # Check for any existing event in the time slot
            if existing_events:
                conflict = True
                conflicting_event = existing_events[0]  # Get the first conflicting event

            if conflict:
                event_start = conflicting_event['start'].get('dateTime', conflicting_event['start'].get('date'))
                dispatcher.utter_message(
                    text=f"Cannot create new event because the time slot is already taken by '{conflicting_event['summary']}' at {event_start}."
                )
            else:
                # No conflict, proceed to add the event
                created_event = add_event(service, event_name, event_start_time, event_end_time)
                if created_event:
                    dispatcher.utter_message(
                        text=f"Event '{event_name}' successfully added to your calendar from {event_start_time.strftime('%d/%m/%y %H:%M:%S')} to {event_end_time.strftime('%d/%m/%y %H:%M:%S')}."
                    )
                else:
                    dispatcher.utter_message(text="Failed to add the event due to an internal error.")

        except ValueError as ve:
            dispatcher.utter_message(text=f"Time parsing error: {ve}. Please ensure the time is in 'DD/MM/YY HH:MM:SS' format.")
        except Exception as e:
            # Handle unexpected errors
            dispatcher.utter_message(text=f"An unexpected error occurred: {str(e)}")

        # Reset slots after execution
        return [AllSlotsReset()]

class GetEvent(Action):

    def name(self) -> Text:
        return "action_get_event"

    def run(self, dispatcher: CollectingDispatcher,
            tracker: Tracker,
            domain: Dict[Text, Any]) -> List[Dict[Text, Any]]:

        try:
            service = get_calendar_service()
            if not service:
                dispatcher.utter_message(text="Failed to initialize Google Calendar service.")
                return []

            now = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
            # Fetch events for the next year to ensure coverage
            events = get_events(service, now, (datetime.utcnow() + timedelta(days=365)).isoformat() + 'Z')

            if not events:
                dispatcher.utter_message(text="You have no upcoming events.")
                return []

            messages = []
            for event in events:
                start = event['start'].get('dateTime', event['start'].get('date'))
                summary = event.get('summary', 'No Title')
                messages.append(f"{start}: {summary}")

            events_message = "\n".join(messages)
            dispatcher.utter_message(text=f"Your upcoming events:\n{events_message}")

        except Exception as e:
            dispatcher.utter_message(text=f"An error occurred while fetching events: {str(e)}")

        return []
