from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

class generateBirthdayEvents(object):
    """
    Generated bday events in calendar
    """
    def __init__(self, calendar_name):
        credentials = get_credentials()
        http = credentials.authorize(httplib2.Http())
        self.service = discovery.build('calendar', 'v3', http=http)
        self.birthday_list = []
        
        calendar_list = self.get_calendar_list()
        for calendar_list_entry in calendar_list['items']:
            if calendar_list_entry['summary'] == calendar_name:
               self.calendar_id = calendar_list_entry['id']
               break
               
        
    def get_calendar_list(self):
        return self.service.calendarList().list(pageToken=None).execute()
        
    def bday_event_exists(self, datetime_start, summary):
        """
        Checks if birthday event is in the calendar already
        """

        bday_start = datetime_start.isoformat() + 'Z'
        bday_end = (datetime_start+datetime.timedelta(hours=23, minutes=59)).isoformat() + 'Z'

        eventsResult = self.service.events().list(
            calendarId=self.calendar_id, timeMin=bday_start, timeMax=bday_end, singleEvents=True,
            orderBy='startTime').execute()
        events = eventsResult.get('items', [])

        for event in events:
            if summary == event['summary']:
                return True
        return False
        
    def add_event_to_calendar(self, datetime_start, summary):
        """
        Adds event to calendar
        """
        if not self.bday_event_exists(datetime_start, summary):
            event = {
              'summary': summary,
              'description': summary,
              'start': {
                'date': datetime_start.strftime("%Y-%m-%d"),
              },
              'end': {
                'date': datetime_start.strftime("%Y-%m-%d"),
              },
              'reminders': {
                'useDefault': False,
                'overrides': [
                  {'method': 'email', 'minutes': 24 * 60},
                  {'method': 'popup', 'minutes': 24 * 60},
                ],
              },
            }
            return self.service.events().insert(calendarId=self.calendar_id, body=event).execute()

    def remove_events_from_calendar(self, name, 
                                    bday_start=datetime.datetime(2015,1,1),
                                    bday_end=datetime.datetime(2053,1,1)
                                    ):
        """
        Remove events from calendar
        """
        bday_start = bday_start.isoformat() + 'Z'
        bday_end = bday_end.isoformat() + 'Z'

        events = self.service.events().list(
            calendarId=self.calendar_id,
            q=name,
            timeMin=bday_start, 
            timeMax=bday_end, 
            singleEvents=True,
            orderBy='startTime').execute()

        for event in events['items']:
            if event['summary'].startswith(name):
                self.service.events().delete(calendarId=self.calendar_id, eventId=event['id']).execute()
        return
        
    def add_birthday_event_to_calendar(self, name, birthday_datetime, year_to_add_to):
        """
        Adds a birthday event to the calendar
        """
        birth_year = birthday_datetime.year
        next_age = str(year_to_add_to - birth_year)
        
        age_append_str = "th"
        if next_age.endswith("1") and not next_age.endswith("11"):
            age_append_str = "st"
        elif next_age.endswith("2") and not next_age.endswith("12"):
            age_append_str = "nd"
        elif next_age.endswith("3") and not next_age.endswith("13"):
            age_append_str = "rd"
            
        summary = "{0}'s {1}{2} Birthday!".format(name, next_age, age_append_str)
        self.add_event_to_calendar(datetime.datetime(year_to_add_to, 
                                                     birthday_datetime.month, 
                                                     birthday_datetime.day),
                                   summary)
                                   
    def import_birthdays(self, file_path):
        """
        Import bdays to list
        """
        self.birthday_list = []
        with open(file_path) as bday_file:
            for line in bday_file:
                line = line.strip()
                if line:
                    line_info = line.split(",")
                    self.birthday_list.append({'last':line_info[0].strip(),
                                                'first': line_info[1].strip(),
                                                'bday_datetime': datetime.datetime.strptime(line_info[2].strip(), "%Y-%m-%d")
                                              })

    def add_all_birthdays_to_calendar(self, year_to_add_to):
        
        for person in self.birthday_list:
            self.add_birthday_event_to_calendar("{0} {1}".format(person["first"], person["last"]),
                                                person["bday_datetime"],
                                                year_to_add_to)
def main():
    """
    Creates a birthday reminder for each person for each year in the range.
    """
    bd = generateBirthdayEvents('Snow Birthday Calendar')
    bd.import_birthdays("example_family_bday.csv")
    for i in range(2016, 2051):
        bd.add_all_birthdays_to_calendar(i)
    #bd.remove_events_from_calendar("Peter Snow Snow")



if __name__ == '__main__':
    main()