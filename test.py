import win32com.client
import os
import time
import datetime as dt
# this is set to the current time
date_time = dt.datetime.now()

# This is set to one minute ago; you can change timedelta's argument to whatever you want it to be
last30MinuteDateTime = dt.datetime.now() - dt.timedelta(days=8)

outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
inbox = outlook.GetDefaultFolder(6)

# retrieve all emails in the inbox, then sort them from most recently received to oldest (False will give you the reverse). Not strictly necessary, but good to know if order matters for your search
messages = inbox.Items  # .Restrict("[Unread]=true")
messages.Sort("[ReceivedTime]", True)


last30MinuteMessages = messages.Restrict(
    "[ReceivedTime] >= '" + last30MinuteDateTime.strftime('%m/%d/%Y %H:%M %p')+"'")

print("Current time: "+date_time.strftime('%m/%d/%Y %H:%M %p'))

print("Messages from the past 30 minute:")
c = 0
for message in last30MinuteMessages:
    print(message.subject)
    print(message.ReceivedTime)
    c = c+1

print("The count of meesgaes unread from past 30 minutes ==", c)
