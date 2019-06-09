# avoid ssl error
import urllib3
urllib3.disable_warnings()

# import python-o365
import sys
sys.path.insert(0, './python-o365')

import datetime as dt
from O365 import Account
import password_list

# outlook
client_id = password_list.client_id
client_secret = password_list.client_secret
account = Account((client_id, client_secret))

if not account.is_authenticated:
    account.authenticate(scopes=['basic', 'calendar_all'])

schedule = account.schedule()
calendar = schedule.list_calendars()

print(calendar)

