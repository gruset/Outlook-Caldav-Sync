import datetime as dt

## Generic variables
debug = False # Set to true for more detaild logging
LogToFile = True # Set to true for logging to file
START_TIME = dt.date.today() # Use today as starting date
END_TIME = START_TIME + dt.timedelta(days=365) # Add a 365 days to starting date

## O365 Variables
CLIENT_ID = "ClientId"
CLIENT_SECRET = "ClientSecret"
TENANT_ID = "TenantId"
DEFAULT_RESOURCE = "default@calendar.com"
O365_FORCE_AUTHENTICATION = False # Set to True to force user authentication. Only used for initial login or if refresh token is expired.

## iCloud Variables
caldav_url = "https://calendar.icloud.com/"
username = "UserName"
password = "SecurePassword"
CALENDAR_NAME = "NameOfCalendar"