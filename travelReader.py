# import imaplib
import re
import email
from email.header import decode_header
import webbrowser
import os
import win32com
import datetime
from datetime import timedelta
import win32com.client
import xlsxwriter
from win32com.client import Dispatch

today = datetime.datetime.today()
twoWeeksAgo = datetime.datetime.today() - timedelta(days=7)
rightNow = datetime.datetime.now()

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts
inbox = outlook.Folders["airbnb@vacationstation.com"]
# airbnbFolder = root_folder.Folders["Rental Requests"].Folders["Airbnb"]
airbnbFolder = inbox.Folders["Inbox"]
messages = airbnbFolder.Items
print(str(twoWeeksAgo))
# messages = airbnbFolder.search(None, "(Unseen SUBJECT RE: Reservation at)")
messages = messages.Restrict("[ReceivedTime] > \'" + twoWeeksAgo.strftime('%m/%d/%Y %H:%M %p') +"\'")
# messages.Sort("[ReceivedTime]", True)

listEmailSS = xlsxwriter.Workbook('F:\\AirBnBSpreadsheets\\'+str(today.month) + '.' + str(today.day) + '.' + str(today.year) + '.'+str(rightNow.hour) + '.'+str(rightNow.minute) + '.'+str(rightNow.second) + '.xlsx')
listEmailSheet = listEmailSS.add_worksheet(str(today.month) + '.' + str(today.day)+'.'+str(today.year))
row = 0
col = 0
for message in messages:
    # if row == 10:
    #     break
    if "airbnb.com" in message.SenderEmailAddress and message.subject.startswith('Pending: Reservation Request at') :
    # if "airbnb.com" in message.SenderEmailAddress and message.subject.startswith('Pending: Reservation Request at') and message.UnRead == True:
        print(str(message.Subject))

        propertyName = str(str(message.Subject).split(" for ")[1])
        # .split("at")[1]
        # listEmailSheet.write(row, col, str(message.Subject))
        listEmailSheet.write(row, col, propertyName)
        # col += 1
        listEmailSheet.write(row, col+1, str(message.receivedTime.date()))
        listEmailSheet.write(row, col+2, str(message.body))
        # assert isinstance(message.Forward, object)
        forwardMsg = message.Forward()
        forwardMsg.To = 'mike@vacationstation.com; don@vacationstation.com'
        # forwardMsg.From = 'robot@vacationstation.com'
        # forwardMsg.Body = 'Workbook located in: '+str(listEmailSS.filename) + message.Body
        forwardMsg.Send()
        col = 0
        row += 1
listEmailSS.close()
