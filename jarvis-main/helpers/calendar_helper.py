import sys
import win32com.client
from datetime import datetime, timedelta

def get_outlook_calendar():
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        # 9 refers to the default calendar folder
        return ns.GetDefaultFolder(9)
    except Exception as e:
        print(f"ERROR: Could not access Outlook - {e}")
        return None

def fetch_events(start, end):
    calendar = get_outlook_calendar()
    if not calendar: return
    
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    
    # Filter for events within the time range
    restriction = f"[Start] >= '{start.strftime('%m/%d/%Y %I:%M %p')}' AND [End] <= '{end.strftime('%m/%d/%Y %I:%M %p')}'"
    restricted_items = items.Restrict(restriction)
    
    for item in restricted_items:
        # Format similar to your Swift script: 
        # calName|||title|||timeStr|||endStr|||location|||isAllDay
        start_str = item.Start.strftime("%#I:%M %p")
        end_str = item.End.strftime("%#I:%M %p")
        location = item.Location if item.Location else ""
        is_all_day = "true" if item.AllDayEvent else "false"
        
        # We use "Outlook" as the calendar name by default
        print(f"Outlook|||{item.Subject}|||{start_str}|||{end_str}|||{location}|||{is_all_day}")

def main():
    if len(sys.argv) < 2:
        command = "today"
    else:
        command = sys.argv[1]

    now = datetime.now()

    if command == "today":
        start = now.replace(hour=0, minute=0, second=0)
        end = start + timedelta(days=1)
        fetch_events(start, end)

    elif command == "upcoming":
        hours = int(sys.argv[2]) if len(sys.argv) > 2 else 4
        start = now
        end = now + timedelta(hours=hours)
        fetch_events(start, end)

    elif command == "calendars":
        # In simple Outlook MAPI, we usually just have the main Calendar
        print("Outlook Default Calendar")

    else:
        print("Usage: calendar_helper.py [today|upcoming|calendars]")

if __name__ == "__main__":
    main()