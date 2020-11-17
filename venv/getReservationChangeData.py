from urllib.request import Request, urlopen
import json
import urllib
import smtplib
from email.message import EmailMessage
import jinja2
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
import datetime
from datetime import date



# read from file 'endVersion'
file_object = open("endVersionFile.txt", "a+")
with open('endVersionFile.txt') as f:
    for line in f:
        pass
    last_line = line


# req = Request("https://hsapi.escapia.com/dragomanadapter/hsapi/GetReservationChanges?startVersion="+last_line)
req = Request("https://hsapi.escapia.com/dragomanadapter/hsapi/GetReservationChanges?startVersion=64041481")
req.add_header('Authorization', 'Bearer Y2JmOWRhNTktNTBhNy00YTk2LThlOTItNGMzZWZiZGI1YTRm')
req.add_header('x-homeaway-hasp-api-version', '10')
req.add_header('x-homeaway-hasp-api-endsystem', 'EscapiaVRS')
req.add_header('x-homeaway-hasp-api-pmcid', '2476')

content = urlopen(req).read()
json_obj = json.loads(content)

startVersion = json_obj['startVersion']
endVersion = json_obj['endVersion']
# print('Start Version: ' + str(startVersion))
# print('End Version: ' + str(endVersion))

# 3535, 25 or 587

templateLoader = jinja2.FileSystemLoader(searchpath=".")
templateEnv = jinja2.Environment(loader=templateLoader)
# print(templateEnv)
s = smtplib.SMTP(host='maila4.newtekwebhosting.com', port=587)
s.login('robot@vacationstation.com', 'Bike!0857')
guestDict = {}
for resChange in json_obj['changes']:
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "WELCOME TO LAKE TAHOE HOORAY"
    msg['From'] = "welcome@vacationstation.com"
    msg['To'] = "robot@vacationstation.com"
    resNum = str(resChange['nativePMSID'])
    # print(resNum)

    # guestRequest = Request("https://hsapi.escapia.com/dragomanadapter/hsapi/GetReservationById?id=11327452")
    resReq = Request("https://hsapi.escapia.com/dragomanadapter/hsapi/GetReservationById?id="+resNum)
    resReq.add_header('Authorization', 'Bearer Y2JmOWRhNTktNTBhNy00YTk2LThlOTItNGMzZWZiZGI1YTRm')
    resReq.add_header('x-homeaway-hasp-api-version', '10')
    resReq.add_header('x-homeaway-hasp-api-endsystem', 'EscapiaVRS')
    resReq.add_header('x-homeaway-hasp-api-pmcid', '2476')

    resContent = urlopen(resReq).read()
    resReqJSON = json.loads(resContent)

    for guest in resReqJSON['guests']:
        if guest['isPrimaryGuest'] == True:
            guestDict = guest
    # print("resNum: "+ resNum)
    templ = templateEnv.get_template("test.j2")

    htmlMessage = templ.render(resNum=resNum, guest=guestDict, reservationDates=resReqJSON["stayDateRange"],
                               stayLen=abs(datetime.datetime.strptime(resReqJSON["stayDateRange"]['endDate'], '%Y-%m-%d') -
                                           datetime.datetime.strptime(resReqJSON["stayDateRange"]['startDate'], '%Y-%m-%d')).days)
    # print(htmlMessage)
    # html = templ.render(resNum=str(resNum))
    # body = html.encode("utf8")
    msg.attach(MIMEText(htmlMessage, 'html'))
    # print(resReqJSON)
    print(guestDict['addresses'][0]['street1'])

    # IMPORTANT
    s.send_message(msg)
s.quit()



# store endVersion to file
file_object.write("\n"+str(endVersion))
file_object.close()

# perform additional operations
