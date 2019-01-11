#!/usr/bin/env python3
import csv
import imaplib
import re
import time
import smtplib
import imapy
import getpass
from datetime import datetime
from email import encoders
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.utils import formatdate
from distutils.util import strtobool
TODAY = datetime.today().strftime('%Y-%m-%d')
TIMESTAMP = str(int(time.time()))

#########################################CONFIG################################################
LOGIN = "LOGIN"
#PASS = getpass.getpass('Enter pass for %s' % LOGIN)
PASS = "PASS"
IMAP = "IMAP SERVER"
SMTP = "SMTP SERVER"
SUBJECT = "SUBJECT" + TODAY
CC = "CC"
TO = "TO" #receiver e-mail
FROM = LOGIN + "@MAILBOX" #your e-mail
FOLDER = 'FOLDER' #folder name like in mailbox (NO RUSSIAN SYMBOLS)
CSVFILE = "new_offer_" + TODAY + '_' + TIMESTAMP + ".csv"
###############################################################################################

def writeFile(filename, mail_list):
    with open(filename, "w+") as output:
        writer = csv.writer(output, lineterminator='\n')
        for val in mail_list:
            writer.writerow([val])

def getOfferUsers2(host,user,ps,folder_name):
    global HAS_NEW_EMAILS

    offer_mail_list = []

    mail = imapy.connect(host = host,
                 username = user,
                 password = ps,
                 ssl= True)

    emails = mail.folder(folder_name).emails()

    if emails:
        HAS_NEW_EMAILS = True
        for em in emails:
            emString = em['text'][0]['text_normalized'] #get mail body as string
            reg = re.search(r'>.*@.*<', emString)
            t = reg.group(0)
            offerEmail = t[1:len(t)-1] # delete > and <
            offer_mail_list.append(offerEmail) #push to list

        writeFile(CSVFILE,offer_mail_list)
        #if strtobool(input('Move mails to Trash?')): #asking if we want to move mails to trash
        for em in emails:
            em.move('offer-processed')
    else:
        HAS_NEW_EMAILS = False

    mail.logout()

def sendMail(host, user, ps, sender, subject, to, filename, cc):
    server = smtplib.SMTP(host, 25)
   # server.login(user,ps)
    header = 'Content-Disposition', 'attachment; filename="%s"' % filename
    attachment = MIMEBase('application', "octet-stream")
    msg = MIMEMultipart()
    msg["From"] = sender
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["To"] = to
    msg["cc"] = cc

    with open(filename, "rb") as fh:
        data = fh.read()
    attachment.set_payload( data )
    encoders.encode_base64(attachment)
    attachment.add_header(*header)
    receivers = [cc, to]
    msg.attach(attachment)
    server.sendmail(sender, receivers, msg.as_string())
    server.quit()

getOfferUsers2(IMAP,LOGIN,PASS,FOLDER)

if HAS_NEW_EMAILS:
    sendMail(SMTP,LOGIN,PASS,FROM,SUBJECT,TO,CSVFILE, CC)
else:
    print ('No new emails')