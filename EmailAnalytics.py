import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import win32com
import win32com.client
import xlsxwriter

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts
inbox = outlook.GetDefaultFolder(6)
print(accounts[0])
messages = inbox.Items
listEmailSS = xlsxwriter.Workbook('C:\\Users\\Don\\Desktop\\emailList.xlsx')
listEmailSheet = listEmailSS.add_worksheet('temp')
row = 0
col = 0
for message in messages:
    if "escapia.com" in message.SenderEmailAddress:
        print(str(message.Subject))
        print(str(message.ReceivedTime))

        listEmailSheet.write(row, col, str(message.Subject))
        row += 1
        if row > 1:
            break
listEmailSS.close()

