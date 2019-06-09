# temp: to avoid ssl error
import urllib3
urllib3.disable_warnings()

import datetime
from slackclient import SlackClient
import password_list
import json


# import python-o365
import sys
sys.path.insert(0, './python-o365')
from O365 import Account

# outlook
client_id = password_list.client_id
client_secret = password_list.client_secret
account = Account((client_id, client_secret))
# slack
token = password_list.token
channel = password_list.channel


def getOlcalEventsAllday(account,calName,calDay):
  if not account.is_authenticated:
      account.authenticate(scopes=['basic', 'calendar_all'])

  schedule = account.schedule()
  calendar = schedule.list_calendars()

  calendar = schedule.get_calendar(calendar_name=calName)
  q = calendar.new_query('start').greater_equal(calDay + ' 00:00:00')
  q.chain('and').on_attribute('end').less_equal(calDay + ' 23:59:59')

  cal = calendar.get_events(query=q, order_by='end/dateTime', include_recurring=True)  # include_recurring=True will include repeated events on the result set.

  return cal


def postSlack(token, channel, msg):
  slack_token = token
  sc = SlackClient(slack_token)
  sc.api_call(
    "chat.postMessage",
    channel=channel,
    #text="<@channel>" + msg1,
    text=msg,
    as_user=True
  )


def postSlackOlEventsAllday(calDay, account, calName, token, channel):
  msg = '`' + calDay + '` \n```\n'
  for l in getOlcalEventsAllday(account,calName,calDay):
    msg = msg + str(l).replace('Subject: ', '* ') + '\n'

  msg = msg + '```'
  
  # temp
  print(msg)
  
  postSlack(token, channel, msg)



if __name__ == "__main__":
  # get Events
  calName = 'Calendar'
  calNameFmt = '*' + calName + '*'
  postSlack(token, channel, calNameFmt)

  ## get today events
  calDay = str(datetime.date.today()) #'%Y-%m-%d'
  postSlackOlEventsAllday(calDay, account, calName, token, channel)

  ## get next business day events
  if datetime.date.today().weekday() <  4:
    calDay2 = str(datetime.date.today() + datetime.timedelta(days=1))
  elif datetime.date.today().weekday() == 4:
    calDay2 = str(datetime.date.today() + datetime.timedelta(days=3))
  elif datetime.date.today().weekday() == 5:
    calDay2 = str(datetime.date.today() + datetime.timedelta(days=2))
  else:
    calDay2 = str(datetime.date.today() + datetime.timedelta(days=1))
  postSlackOlEventsAllday(calDay2, account, calName, token, channel)

  # get tasks
  calName2 = 'Tasks'
  calNameFmt2 = '*' + calName2 + '*'
  postSlack(token, channel, calNameFmt2)

  ## get today tasks
  postSlackOlEventsAllday(calDay, account, calName2, token, channel)

  ## get next business day tasks
  postSlackOlEventsAllday(calDay2, account, calName2, token, channel)

