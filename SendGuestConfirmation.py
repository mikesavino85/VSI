import datetime
from datetime import timedelta
import json
import sqlalchemy
from flask import jsonify
from sqlalchemy import create_engine
import jinja2
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
import pdfkit

options = {
  "enable-local-file-access": None
}
SQLEngine = create_engine("mssql+pyodbc://mike:in7ju1Ohy3@35.155.42.165:1433/VSIWebsite_AWS?driver=SQL+Server+Native+Client+11.0")

templateLoader = jinja2.FileSystemLoader(searchpath="./Templates/GuestConfirmation/")
templateEnv = jinja2.Environment(loader=templateLoader)

s = smtplib.SMTP(host='maila4.newtekwebhosting.com', port=587)
s.login('robot@vacationstation.com', 'Bike!0280')


conn = SQLEngine.connect()
quoteTable = conn.execute("SELECT q.booktype, q.status, q.quotenum, q.custcode, convert(varchar, q.rstartdate, 101) startDate, datename(dw, q.rstartDate) startDay, datename(dw, q.renddate) endDay,   convert(varchar, q.renddate, 101) enddate, q.numofdays,   q.prop_id, p.numofsleeps, p.address1, p.security, convert(varchar,convert(decimal(8,2),q.gtotal)) gtotal,   convert(varchar,convert(decimal(8,2),q.ltaxvalu)) ltaxvalu,   convert(varchar,convert(decimal(8,2), q.ext_goods)) ext_goods,    convert(varchar,convert(decimal(8,2), q.ext_tax)) ext_tax, convert(varchar,convert(decimal(8,2), q.trav_amt)) trav_amt, q.DateCanc, q.custtype, p.areacode,   q.lonorsea, net.recdescr wifi, pw.recdescr pw,  c.name, isNull(c.address1, '') custAddress1, isNull(c.address2, '') custAddress2, isnull(c.city, '') custCity, isNull(nullif(c.state, '.'), '') custState, isNull(c.zip,'') custZip, isnull(c.phone, '') custPhone, isNull(c.email, '') custEmail, isnull(c.CellPhone, '') custCellPhone, convert(varchar, GETDATE(), 101) todayDate, convert(varchar,convert(decimal(8,2), q.ltaxvalu + q.ext_tax)) countyroomtax, convert(varchar,convert(decimal(8,2), q.ext_goods - 74.00)) cleanfee, convert(varchar,convert(decimal(8,2), 74.00)) resfee, convert(varchar,convert(decimal(8,2), q.ltaxvalu + q.ext_tax + q.ext_goods + q.gtotal)) totalprice, convert(varchar,p.latitude) latitude, convert(varchar,p.longitude) longitude, pd.parkingLimit FROM  [24.176.189.162, 1433\S1372].[S1372].propertyplus.quote q left join [24.176.189.162, 1433\S1372].[S1372].propertyplus.prope p on q.prop_id = p.prop_id left join (select recdescr, prop_id from [24.176.189.162, 1433\S1372].[S1372].propertyplus.propwarr where recdescr is not null and recSeq = 3) net on net.prop_id = p.prop_id left join (select recdescr, prop_id from [24.176.189.162, 1433\S1372].[S1372].propertyplus.propwarr where recdescr is not null and recSeq = 4) pw on pw.prop_id = p.prop_id left join [24.176.189.162, 1433\S1372].[S1372].propertyplus.custc c on c.custcode = q.custcode left join propertyData pd on pd.propertyCode = p.prop_id   WHERE ((q.rstartdate>=GetDate())) and q.lonorsea = 'S' and q.booktype in ('CONF', 'AIRB') ORDER BY q.rstartdate, q.prop_id "
                          )

resInfo = json.dumps([dict(r) for r in quoteTable], default=str)

resInfoJSON = json.loads(resInfo)

count = 1
for resData in resInfoJSON:

    msg = MIMEMultipart('html')
    msg['Subject'] = "Lake Tahoe Vacation Rental | Vacation Station (775) 831-3664 | "+str(resData['prop_id'])+" "+str(resData['quotenum'])
    msg['From'] = "robot@vacationstation.com"
    msg['To'] = "mikesavino85@gmail.com, robot@vacationstation.com, mike@vacationstation.com, kathleen.savino@gmail.com, neda@vacationstation.com"

    payTable = conn.execute("select convert(varchar, pay_date, 101) pay_date, convert(varchar,convert(decimal(8,2), payvalue)) payvalue, pay_meth, cc_type from [24.176.189.162, 1433\S1372].[S1372].propertyplus.qtoph where quotenum = ?", resData['quotenum'])
    payInfo = json.dumps([dict(r) for r in payTable], default=str)
    payInfoJSON = json.loads(payInfo)

    notificationTable = conn.execute("select propertyNotificationCode from propertyNotification where propertyCode = ?", resData['prop_id'])
    notificationInfo = json.dumps([dict(r) for r in notificationTable], default=str)
    notificationInfoJSON = json.loads(notificationInfo)

    securityTable = conn.execute("select pd.propertyCode, pst.securityTypeDescription, count(pst.securityTypeDescription) countItem from propertyData pd left join propertySecurity ps on pd.propertyID = ps.propertyID left join propertySecurityType pst on ps.securityTypeID = pst.securityTypeID group by propertyCode, pst.securityTypeDescription having count(pst.securityTypeDescription) > 0 and pd.propertyCode = ?", resData['prop_id'])
    securityInfo = json.dumps([dict(r) for r in securityTable], default=str)
    securityInfoJSON = json.loads(securityInfo)

    templ = templateEnv.get_template("GuestRegistration.j2")
    templPDF = templateEnv.get_template("GuestRegistrationSignature.j2")
    htmlMessage = templ.render(resData=resData, payInfo=payInfoJSON, notificationInfo=notificationInfoJSON)
    pdfMessage = templPDF.render(resData=resData, securityInfo=securityInfoJSON)
    msg.attach(MIMEText(htmlMessage, 'html'))
    pdfkit.from_string(pdfMessage, 'test.pdf', options=options)
    with open ('test.pdf', "rb") as testpdf:
        testpdfopened = testpdf.read()
    attachedfile = MIMEApplication(testpdfopened, _subtype = "pdf")
    attachedfile.add_header('content-disposition', 'attachment', filename = str(resData['prop_id'])+"_"+str(resData['quotenum'])+"_Signature.pdf")
    msg.attach(attachedfile)
    s.send_message(msg)
    count = count + 1
    if count == 11:
        exit()

exit()
