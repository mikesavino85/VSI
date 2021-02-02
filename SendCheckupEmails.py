import re
import email
from dataclasses import dataclass
from email.header import decode_header
import webbrowser
import os
import win32com
import datetime
from datetime import timedelta
import win32com.client
import xlsxwriter
from win32com.client import Dispatch
import pyodbc
import json
import sqlalchemy
from flask import jsonify
from sqlalchemy import create_engine
import jinja2
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr

SQLEngine = create_engine("mssql+pyodbc://PropertyPlus:propertyplus@192.168.1.10:1433/S1372?driver=SQL+Server+Native+Client+11.0")

today = datetime.datetime.today()
twoWeeksAgo = datetime.datetime.today() - timedelta(days=7)
rightNow = datetime.datetime.now()

templateLoader = jinja2.FileSystemLoader(searchpath="./Templates")
templateEnv = jinja2.Environment(loader=templateLoader)

s = smtplib.SMTP(host='maila4.newtekwebhosting.com', port=587)
# s.login('robot@vacationstation.com', 'Bike!0857')
s.login('robot@vacationstation.com', 'Bike!0280')


conn = SQLEngine.connect()
quoteTable = conn.execute("select top 1 convert(varchar(8000), rstartDate, 1) startDate, "
                          "convert(varchar(8000), renddate, 1) enddate, "
                          "q.custcode custCode, "
                          "q.name, p.prop_id propcode, c.email  "
                          "from propertyplus.quote q "
                          "left join propertyPlus.custc c on c.custcode = q.custcode "
                          "left join propertyPlus.prope p on q.prop_id = p.prop_id "
                          "where rstartdate > dateadd(wk, -1, getdate())"
                          "     and rstartdate <= getdate()"
                          "     and renddate >= getdate()"
                          "     and q.custtype in ('IB', 'vsweb', 'sumfest', 'vrbo', 'airbnb', 'repeat', 'walkin', 'word', 'resched')")

# print(json.dumps([dict(r) for r in quoteTable]))

resInfo = json.dumps([dict(r) for r in quoteTable])
resInfoJSON = json.loads(resInfo)
# resInfoJSON = jsonify(quoteTable)
# print(resInfoJSON[1])

for guestContact in resInfoJSON:
    # print(guestContact['startDate'])
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "Hello, thank you for staying with Vacation Station"
    msg['From'] = "robot@vacationstation.com"
    msg['To'] = "mikesavino85@gmail.com, mike@vacationstation.com, robot@vacationstation.com"

    templ = templateEnv.get_template("confirmation1.j2")
    # templ = templateEnv.get_template("baseEmail.j2")

    htmlMessage = templ.render(guestContact=guestContact)
    msg.attach(MIMEText(htmlMessage, 'html'))

    # print(htmlMessage)
    # html = templ.render(resNum=str(resNum))
    # body = html.encode("utf8")
    s.send_message(msg)
exit()
