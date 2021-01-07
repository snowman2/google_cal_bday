import datetime
import pickle
import os.path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.pickle.
SCOPES = ["https://www.googleapis.com/auth/calendar"]
CLIENT_SECRET_FILE = "credentials.json"
APPLICATION_NAME = "Birthday Calendar Generator"


def get_credentials():
    # From: https://developers.google.com/calendar/quickstart/python
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return creds


class GenerateBirthdayEvents:
    """
    Generated bday events in calendar
    """

    def __init__(self, calendar_name):
        self.service = build("calendar", "v3", credentials=get_credentials())
        self.birthday_list = []

        calendar_list = self.service.calendarList().list(pageToken=None).execute()
        for calendar_list_entry in calendar_list["items"]:
            if calendar_list_entry["summary"] == calendar_name:
                self.calendar_id = calendar_list_entry["id"]
                break
        else:
            raise RuntimeError("Calendar not found.")

    def bday_event_exists(self, datetime_start, summary):
        """
        Checks if birthday event is in the calendar already
        """

        bday_start = datetime_start.isoformat() + "Z"
        bday_end = (
            datetime_start + datetime.timedelta(hours=23, minutes=59)
        ).isoformat() + "Z"

        eventsResult = (
            self.service.events()
            .list(
                calendarId=self.calendar_id,
                timeMin=bday_start,
                timeMax=bday_end,
                singleEvents=True,
                orderBy="startTime",
            )
            .execute()
        )
        events = eventsResult.get("items", [])

        for event in events:
            if summary == event["summary"]:
                return True
        return False

    def add_event_to_calendar(self, datetime_start, summary):
        """
        Adds event to calendar
        """
        if not self.bday_event_exists(datetime_start, summary):
            event = {
                "summary": summary,
                "description": summary,
                "start": {
                    "date": datetime_start.strftime("%Y-%m-%d"),
                },
                "end": {
                    "date": datetime_start.strftime("%Y-%m-%d"),
                },
                "reminders": {
                    "useDefault": False,
                    "overrides": [
                        {"method": "email", "minutes": 24 * 60},
                        {"method": "popup", "minutes": 24 * 60},
                    ],
                },
            }
            return (
                self.service.events()
                .insert(calendarId=self.calendar_id, body=event)
                .execute()
            )

    def remove_events_from_calendar(
        self,
        name,
        bday_start=datetime.datetime(2015, 1, 1),
        bday_end=datetime.datetime(2053, 1, 1),
    ):
        """
        Remove events from calendar
        """
        bday_start = bday_start.isoformat() + "Z"
        bday_end = bday_end.isoformat() + "Z"

        events = (
            self.service.events()
            .list(
                calendarId=self.calendar_id,
                q=name,
                timeMin=bday_start,
                timeMax=bday_end,
                singleEvents=True,
                orderBy="startTime",
            )
            .execute()
        )

        for event in events["items"]:
            if event["summary"].startswith(name):
                self.service.events().delete(
                    calendarId=self.calendar_id, eventId=event["id"]
                ).execute()
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
        self.add_event_to_calendar(
            datetime.datetime(
                year_to_add_to, birthday_datetime.month, birthday_datetime.day
            ),
            summary,
        )

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
                    self.birthday_list.append(
                        {
                            "last": line_info[0].strip(),
                            "first": line_info[1].strip(),
                            "bday_datetime": datetime.datetime.strptime(
                                line_info[2].strip(), "%Y-%m-%d"
                            ),
                        }
                    )

    def add_all_birthdays_to_calendar(self, year_to_add_to):

        for person in self.birthday_list:
            self.add_birthday_event_to_calendar(
                "{0} {1}".format(person["first"], person["last"]),
                person["bday_datetime"],
                year_to_add_to,
            )


def main():
    """
    Creates a birthday reminder for each person for each year in the range.
    """
    bd = GenerateBirthdayEvents("Snow Birthday Calendar")
    bd.import_birthdays("example_family_bday.csv")
    for year in range(2021, 2051):
        bd.add_all_birthdays_to_calendar(year)
    # bd.remove_events_from_calendar("Frosty Snowman")


if __name__ == "__main__":
    main()
