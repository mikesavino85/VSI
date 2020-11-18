mport imaplib
import email
from email.header import decode_header
import webbrowser
import os
import win32com
import datetime
import win32com.client
import xlsxwriter
from win32com.client import Dispatch
today = datetime.datetime.today()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts
inbox = outlook.Folders["travel@vacationstation.com"]

root_folder = outlook.Folders.Item(3)
# airbnbFolder = root_folder.Folders["Rental Requests"].Folders["Airbnb"]
airbnbFolder = inbox.Folders["Rental Requests"].Folders["Airbnb"]

messages = airbnbFolder.Items
messages.Sort("[ReceivedTime]", True)
listEmailSS = xlsxwriter.Workbook('C:\\Users\\Don\\Desktop\\AirBnBEmail\\emailList'+str(today.month) + str(today.day)+str(today.year) + '.xlsx')
listEmailSheet = listEmailSS.add_worksheet(str(today.month) + '.'+ str(today.day)+'.'+str(today.year))
row = 0
col = 0
# print (len(messages))
# print (root_folder.Name)
for message in messages:
    # print (message.receivedTime.date())
    # if message.receivedTime.date() < today.date():
    if row == 2:
        break
    if "airbnb.com" in message.SenderEmailAddress and "RE: Reservation at" in message.subject:
        # print(str(message.Subject))
        propertyName = str(str(message.Subject).split(" for ")[0]).split(" at ")[1]
        print(message.body)
        # .split("at")[1]
        # listEmailSheet.write(row, col, str(message.Subject))
        listEmailSheet.write(row, col, propertyName)
        # col += 1
        listEmailSheet.write(row, col+1, str(message.receivedTime.date()))
        listEmailSheet.write(row, col+2, str(message.body))
        col = 0
        row += 1
listEmailSS.close()
