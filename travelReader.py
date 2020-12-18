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
airbnbFolder = inbox.Folders["Inbox"]
messages = airbnbFolder.Items
print(str(twoWeeksAgo))

messages = messages.Restrict("[ReceivedTime] > \'" + twoWeeksAgo.strftime('%m/%d/%Y %H:%M %p') +"\'")

listEmailSS = xlsxwriter.Workbook('F:\\AirBnBSpreadsheets\\'+str(today.month) + '.' + str(today.day) + '.' + str(today.year) + '.'+str(rightNow.hour) + '.'+str(rightNow.minute) + '.'+str(rightNow.second) + '.xlsx')
listEmailSheet = listEmailSS.add_worksheet(str(today.month) + '.' + str(today.day)+'.'+str(today.year))
row = 0
col = 0
for message in messages:
    # if row == 10:
    #     break
    if "airbnb.com" in message.SenderEmailAddress and message.subject.startswith('Pending: Reservation Request at') :

        print(str(message.Subject))

        propertyName = str(str(message.Subject).split(" for ")[1])
        listEmailSheet.write(row, col, propertyName)
        listEmailSheet.write(row, col+1, str(message.receivedTime.date()))
        listEmailSheet.write(row, col+2, str(message.body))
        forwardMsg = message.Forward()
        forwardMsg.To = 'mike@vacationstation.com; don@vacationstation.com'
        forwardMsg.Send()
        col = 0
        row += 1
listEmailSS.close()
