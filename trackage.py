import smtplib
import time
import imaplib
import email
from dateutil import parser
from datetime import datetime, timedelta, date
from dateutil.relativedelta import relativedelta
import re
import requests
import quopri
from colorama import Fore, Style
import sys
import yaml

try:
  file = open('./config.yml')
except:
  exit()

itemsList = yaml.load(file, Loader=yaml.FullLoader)

FROM_EMAIL = itemsList['email']
FROM_PWD = itemsList['password']
SMTP_SERVER = itemsList['email_smtp_server']

SHIPPO_API_TOKEN = itemsList['shippoToken']

UPS_REGEX_STRING = "1Z\w{16}"
USPS_REGEX_STRING = "\d{20,22}"
FEDEX_REGEX_STRING = "\d{12,22}"
DHL_REGEX_STRING = "\d{10,11}"

def get_text(msg):
    if msg.is_multipart():
        return get_text(msg.get_payload(0))
    else:
        return msg.get_payload(None, True)

def dt_parse(t):
    ret = datetime.strptime(t[0:16],'%Y-%m-%dT%H:%M')
    if t[18]=='+':
        ret+=timedelta(hours=int(t[19:22]),minutes=int(t[23:]))
    elif t[18]=='-':
        ret-=timedelta(hours=int(t[19:22]),minutes=int(t[23:]))
    return ret

def find_matches(regexString, carrierName, email_text):
    return re.findall(
        regexString,
        email_text,
        flags=re.IGNORECASE
    )

def fetch_package_status(service, regex_string, text):
    BASE_URL = "https://api.goshippo.com/tracks/"

    matches = find_matches(regex_string, service, text)

    if matches:
        url = BASE_URL + service + "/" + matches[0]
        
        response = requests.get(url,headers={"Authorization": SHIPPO_API_TOKEN})

        response_json = response.json()

        return response_json

def tracking_url(service, tracking_number):
    switcher = {
        "ups": "https://www.ups.com/track?loc=en_US&tracknum=" + tracking_number + "&requester=WT/",
        "usps": "https://tools.usps.com/go/TrackConfirmAction?qtc_tLabels1=" + tracking_number,
        "fedex": "https://www.fedex.com/apps/fedextrack/?tracknumbers=" + tracking_number,
        "dhl_ecommerce": "https://www.dhl.com/en/express/tracking.html?brand=DHL&AWB=" + tracking_number
    }

    return switcher.get(service, "") 

def colorForStatus(status):
    switcher = {
        "PRE_TRANSIT": Fore.YELLOW,
        "TRANSIT": Fore.YELLOW,
        "DELIVERED": Fore.GREEN,
        "RETURNED": Fore.GREEN,
        "FAILURE": Fore.RED,
        "UNKNOWN": Fore.RED
    }

    return switcher.get(status, "")

def read_email():

    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        mail.select('inbox')

        today = datetime.today()
        one_month_ago_date = date.today() + relativedelta(months=-1)
        
        type, data = mail.search(None, "(SINCE " + one_month_ago_date.strftime("%d-%b-%Y") + " OR (TEXT \"Track\") (TEXT \"Tracking\"))")
        mail_ids = data[0]
        id_list = map(lambda x: int(x), mail_ids.split())

        printed_id_set = set()

        for i in reversed(id_list):

            typ, data = mail.fetch(i, '(RFC822)' )

            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(response_part[1])
                    
                    email_text = get_text(msg)

                    ups_status = fetch_package_status("ups", UPS_REGEX_STRING, email_text)
                    usps_status = fetch_package_status("usps", USPS_REGEX_STRING, email_text)
                    fedex_status = fetch_package_status("fedex", FEDEX_REGEX_STRING, email_text)
                    dhl_status = fetch_package_status("dhl_ecommerce", DHL_REGEX_STRING, email_text)

                    status_list = [ups_status, usps_status, fedex_status, dhl_status]

                    for i, val in enumerate(status_list): 
                        if val and val["tracking_status"] and val["tracking_number"] not in printed_id_set:
                            status = val["tracking_status"]["status"]
                            sys.stdout.write("Found tracking number in email with subject: \"" + msg["Subject"] + "\"\n")
                            
                            sys.stdout.write("Status: " + colorForStatus(status) + status + "\n")
                            sys.stdout.write(Style.RESET_ALL)
                            
                            if status == "TRANSIT" or status == "PRE_TRANSIT":
                                date_string = datetime.strptime(val["eta"], "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%a, %B %d")
                                sys.stdout.write("ETA: " + date_string + "\n")
                            
                            sys.stdout.write("Carrier: " + val["carrier"] + "\n")
                            sys.stdout.write("Tracking number: " + val["tracking_number"] + "\n")
                            sys.stdout.write(tracking_url(val["carrier"], val["tracking_number"]) + "\n\n")

                            printed_id_set.add(val["tracking_number"])

    except Exception, e:
        print str(e)
        exit()

read_email()
