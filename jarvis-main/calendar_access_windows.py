import win32com.client
import datetime

class OutlookCalendar:
    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        self.calendar = self.namespace.GetDefaultFolder(9)  # 9 refers to the calendar folder
        self.cache = []

    def fetch_events(self, start_date, end_date):
        self.cache = []  # Clear cache
        try:
            items = self.calendar.Items
            items.IncludeRecurrences = True
            items.Sort('[Start]')
            items = items.Restrict(f"[Start] >= '{start_date}' AND [End] <= '{end_date}'")

            for item in items:
                self.cache.append({
                    'subject': item.Subject,
                    'start': item.Start,
                    'end': item.End,
                    'location': item.Location
                })
        except Exception as e:
            print(f'Error fetching events: {e}')  

    def format_events_for_voice(self):
        formatted_events = []
        for event in self.cache:
            start_time = event['start'].strftime('%Y-%m-%d %H:%M')
            formatted_events.append(f'Event: {event['subject']} at {start_time} in {event['location']}')
        return '\n'.join(formatted_events)

    def fetch_and_format(self, days_ahead=7):
        today = datetime.datetime.now()
        start_date = today.strftime('%Y-%m-%d')
        end_date = (today + datetime.timedelta(days=days_ahead)).strftime('%Y-%m-%d')
        self.fetch_events(start_date, end_date)
        return self.format_events_for_voice()
