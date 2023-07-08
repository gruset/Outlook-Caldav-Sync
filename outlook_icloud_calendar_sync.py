import datetime as dt
from O365 import Account, MSGraphProtocol, Connection
import caldav
import config_user as conf
from icalendar import Calendar
import json
import readline

## Generic Variables
START_TIME = conf.START_TIME
END_TIME = conf.END_TIME
debug = conf.debug

## O365 Variables
CLIENT_ID = conf.CLIENT_ID
CLIENT_SECRET = conf.CLIENT_SECRET
TENANT_ID = conf.TENANT_ID

## iCloud Variables
caldav_url = conf.caldav_url
username = conf.username
password = conf.password
CALENDAR_NAME = conf.CALENDAR_NAME

## Generic Functions
def LogToConsole(message):
    logString = "[" + dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + "] > " + str(message)
    
    if conf.LogToFile:
        try:
            f = open("sync_log.txt", "a")
            f.write(logString + '\n')
            f.close()
        except Exception as e:
            print(e)

    print(logString)

def RotateLogFile(event):
    if event == "start":
        return
    if event == "succes":
        return

### Main logic ###
"""
First we retrieve all of the events (including recurring) from O365 in the selected period.
These are added to a list with the specified attributes in a class.

Then we retrieve all of the events from iCloud in the selected period.
For each event we check if it exists in the corresponding list from O365, based on
    1. summary
    2. starttime
    3. endtime
if it does not exist in O365 it is delted. This also happens if the event "exists" but has varying attributes, and just needs to be updated. It is easier to just delte and create later.
If the event exists it is addet to a list with the specified attributes in a class.

After the classes have been populated, they are compared with len() to see if there are objects in O365 that does not exist in iCloud.
If the length of classes are different, the O365 class is traversed in order to find events that are missing in iCloud. These are then created.

That's all folks!
"""
RotateLogFile("start") # First - rotate the log file

LogToConsole("==================")
LogToConsole("Beginning new sync")
LogToConsole("==================")

LogToConsole("Getting Calendar events from " + str(START_TIME) + " to " + str(END_TIME))

## O365 stuff
credentials = (CLIENT_ID, CLIENT_SECRET)
protocol = MSGraphProtocol(default_resource=conf.DEFAULT_RESOURCE)
account = Account(credentials, tenant_id=TENANT_ID, protocol=protocol)

### Authentication logic ###
"""
The auth logic is outcommented. Only required for initial authentication on "delegated" permissions.
We use the refresh token to maintain the connection.
"""
# scopes = [ 'basic', 'calendar' ]
# if account.authenticate(scopes=scopes):
#     LogToConsole('O365 Authenticated')
# else:
#     raise Exception("O365 Authentication Error")

schedule = account.schedule()
calendar = schedule.get_default_calendar()

q = calendar.new_query('start').greater_equal(START_TIME)
q.chain('and').on_attribute('end').less_equal(END_TIME)
events = calendar.get_events(query=q, include_recurring=True)

# Create a class to hold O365 Event data
class O365Event:
    def __init__(self, subject, start, end, description, location, uid = 1):
        self.subject = subject
        self.start = start
        self.end = end
        self.description = description
        self.location = location
        self.uid = uid

O365Events = []
for event in events:
    O365Events.append(
        O365Event(
            event.subject, 
            event.start, 
            event.end, 
            event.body,
            event.location
        )
    )
# Count the number of events in the list
eventCountO365 = len(O365Events)

if debug:
    LogToConsole("There are " + str(eventCountO365) + " events in the period")

## iCloud stuff
client = caldav.DAVClient(url=caldav_url, username=username, password=password)
my_principal = client.principal()
LogToConsole("iCloud Authenticated")

try:
    calendars = my_principal.calendars()
    my_calendar = my_principal.calendar(name=CALENDAR_NAME)
    if debug:
        LogToConsole("Selected calendar: " + my_calendar.name)
except Exception as e:
    raise Exception(e)

# Create a class to hold iCloud Event data
class CalDavEvent:
    def __init__(self, subject, start, end, description, location, uid = 1):
        self.subject = subject
        self.start = start
        self.end = end
        self.description = description
        self.location = location
        self.uid = uid

CalDavEvents = [] 

events_fetched = my_calendar.search(
        start=START_TIME,
        end=END_TIME,
        event=True,
        expand=True
    )

for event in events_fetched:
    obj = event.data
    cal = Calendar.from_ical(obj)
    
    for e in cal.walk('vevent'):
        summary = e.get('summary')
        description = e.get('description')
        dtStart = e.get('dtstart')
        dtEnd = e.get('dtend')
        location = e.get('location')

        """
        If the event does not exist in O365 with:
            1. The same subject
            2. The same starttime
            3. The same endtime
        it will be deleted. It also handles if something of the above is changed. In that case a new event will be created afterwards.
        """
        if not any(x for x in O365Events if x.subject == summary and x.start == dtStart.dt and x.end == dtEnd.dt):
            # Delete event not existing in O365
            LogToConsole("Event with subject [" + summary + "] does not exist in O365. Deleting...")
            event.delete()
        else:
            # If it exists, add it to the comparison list
            CalDavEvents.append(
            CalDavEvent(
                subject=summary,
                start=dtStart.dt,
                end=dtEnd.dt, 
                description=description,
                location=location
            )
        )

eventCountCalDav = len(CalDavEvents)

if debug:
    LogToConsole("There are " + str(eventCountCalDav) + " events in the period")


compareNumberOfEvents = eventCountCalDav - eventCountO365

if compareNumberOfEvents == 0:
    LogToConsole("Same number of events in both calendars. All is fine!")
elif compareNumberOfEvents < 0:
    LogToConsole(str(abs(compareNumberOfEvents)) + " events are missing in iCloud calendar")
    for e in O365Events:
        if not any(x for x in CalDavEvents if x.subject == e.subject and x.start == e.start and x.end == e.end):
            LogToConsole("MISSING - The subject " + e.subject + " DOES NOT exists i iCloud")
            try:
                LogToConsole("Creating event")
                
                try:
                    locationJson = e.location#
                    locationData = json.load(locationJson)
                    location = locationData.displayName
                    LogToConsole("Location data: " + location)
                except:
                    location = ''#e.location

                my_event = my_calendar.save_event(
                    dtstart=e.start,
                    dtend=e.end,
                    summary=e.subject,
                    location=location
                )
            except Exception as e:
                raise Exception(e)
elif compareNumberOfEvents > 0:
    LogToConsole("There are " + str(compareNumberOfEvents) + "more events in iCloud than i O365 and I ended up here without creating anything. Something is wrong.")