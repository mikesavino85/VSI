from urllib.request import Request, urlopen
import json
import urllib
# import smtplib
# from email.message import EmailMessage
# import jinja2
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.header import Header
# from email.utils import formataddr
import datetime
from datetime import date

req = Request("https://hsapi.escapia.com/dragomanadapter/hsapi/GetReservationChanges?startVersion=64041481")
req.add_header('Authorization', 'Bearer Y2JmOWRhNTktNTBhNy00YTk2LThlOTItNGMzZWZiZGI1YTRm')
req.add_header('x-homeaway-hasp-api-version', '10')
req.add_header('x-homeaway-hasp-api-endsystem', 'EscapiaVRS')
req.add_header('x-homeaway-hasp-api-pmcid', '2476')


bcp vacationstation.dbo.rentals in "C:\Users\Don\Desktop\DB Exports\Rentals.txt" -c -T -t -S "SQLB50.newtekwebhosting.com"