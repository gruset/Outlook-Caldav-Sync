import datetime as dt

## Generic variables
debug = False # Set to true for more detaild logging
LogToFile = True # Set to true for logging to file
today = dt.date.today()
#START_TIME = today - dt.timedelta(days=365) # Use 365 days back as starting date
START_TIME = today # Use today as starting date
END_TIME = START_TIME + dt.timedelta(days=365) # Add a 365 days to starting date
#END_TIME = today + dt.timedelta(days=365) # Use today + 365 days as end date

## O365 Variables
CLIENT_ID = "ClientId"
CLIENT_SECRET = "ClientSecret"
TENANT_ID = "TenantId"
DEFAULT_RESOURCE = "default@calendar.com"
AUTHENTICATE_CONSOLE = False # Set this to True to force interactive console authentication

## iCloud Variables
caldav_url = "https://calendar.icloud.com/"
username = "UserName"
password = "SecurePassword"
CALENDAR_NAME = "NameOfCalendar"