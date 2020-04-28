# Import stuff
import subprocess
import pkg_resources
import sys
import pandas
import time
import configparser
import argparse
import smtplib
import os.path
import signal
import contextlib
import logging as log
from tqdm import tqdm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from mako.template import Template
from email.mime.base import MIMEBase
from email import encoders

#Setup Argument Parser and define arguments
parser = argparse.ArgumentParser(description='Send automated emails from CSV file input.')
parser.add_argument("-v", "--verbose", default=False, help="Verbose output", action="store_true")
parser.add_argument("-i", "--input", dest="csvfile", required=True, help="Path to CSV File", metavar="FILE")
args = parser.parse_args()

#Determine if verbose
if args.verbose == True:
    log.basicConfig(format="%(levelname)s: %(message)s", level=log.DEBUG)
    log.info("Verbose output.")
else:
    log.basicConfig(format="%(levelname)s: %(message)s")

#Keyboard interrupt
def keyboardInterruptHandler(signal, frame):
    print("KeyboardInterrupt has been caught. Stopping...")
    pbar.close()
    exit(0)
signal.signal(signal.SIGINT, keyboardInterruptHandler)

# Config
config = configparser.ConfigParser()
config.sections()
config.read("./config.ini")

# EMAIL FUNCTION
def send_mail(row):
    if args.verbose == True:
        pbar.write(f"Setting up Email for: %(name)s" % {'name': row['Name']})
    recipients = [row['Email']]
    msg = MIMEMultipart('alternative')
    msg['From'] = config['smtp']['senderEmail']
    msg['To'] = ", ".join(recipients)

# Select message to use
    if row['Grade'] == 9:
        message = config['msg']['grd9msg']
        text = config['msg']['grd9msg']
        msg['Subject'] = config['subject']['grd9subject']
        if args.verbose == True:
            pbar.write(f"%(name)s's grade: 9" % {'name': row['Name']})
    elif row['Grade'] == 10:
        message = config['msg']['grd10msg']
        text = config['msg']['grd10msg']
        msg['Subject'] = config['subject']['grd10subject']
        if args.verbose == True:
            pbar.write(f"%(name)s's grade: 10" % {'name': row['Name']})
    elif row['Grade'] == 11:
        message = config['msg']['grd11msg']
        text = config['msg']['grd11msg']
        msg['Subject'] = config['subject']['grd11subject']
        if args.verbose == True:
            pbar.write(f"%(name)s's grade: 11" % {'name': row['Name']})
    elif row['Grade'] == 12:
        message = config['msg']['grd12msg']
        text = config['msg']['grd12msg']
        msg['Subject'] = config['subject']['grd12subject']
        if args.verbose == True:
            pbar.write(f"%(name)s's grade: 12" % {'name': row['Name']})
    else:
        message = config['msg']['assumedGrdMsg']
        text = config['msg']['assumedGrdMsg']
        msg['Subject'] = config['subject']['assumedGrdSubject']
        if args.verbose == True:
            pbar.write(f"Assumed %(name)s's grade: 12" % {'name': row['Name']})

# HTML Email Template
    html = Template(filename=config['branding']['EmailHTMLTemplate'])
    html = html.render(message=message,emailtitle=config['branding']['emailtitle'], brandingLogo=config['branding']['brandinglogo'], heroImage=config['branding']['heroimage'], footerschool=config['branding']['footerschool'], phonenumber=config['branding']['phonenumber'], address=config['branding']['address'], weburl1=config['branding']['weburl1'], weburldisplay1=config['branding']['weburldisplay1'], weburl2=config['branding']['weburl2'], weburldisplay2=config['branding']['weburldisplay2'])

# MIME Conversion and Attach HTML and TXT to Email
    text = MIMEText(text, 'plain')
    html = MIMEText(html, 'html')
    msg.attach(text)
    msg.attach(html)

# Attach Files
    list = row['File'].split(",")
    if args.verbose == True:
        pbar.write(f"File list: %(list)s" % {'list': list})
    for i in list:
        attachment = open("./%s" % i, "rb")
        if args.verbose == True:
            pbar.write(f"Attaching file: " + i)
        attach = MIMEBase('application', 'octet-stream')
        attach.set_payload((attachment).read())
        encoders.encode_base64(attach)
        attach.add_header('Content-Disposition', "attachment; filename= %s" % i)
        msg.attach(attach)

# Setup SMTP Mailer and Send
    mailer = smtplib.SMTP(config['smtp']['emailSMTP'], config['smtp']['emailPORT'])
    mailer.ehlo()
    mailer.starttls()
    mailer.login(config['smtp']['emailLogin'], config['smtp']['emailPassword'])
    mailer.sendmail(config['smtp']['senderEmail'], recipients, msg.as_string())
    mailer.quit()
    if args.verbose == True:
        pbar.write(f"Sent Email to: %(name)s \n" % {'name': row['Name']})

def yes_or_no(question):
    answer = input(question).lower().strip()
    while not(answer == "y" or answer == "yes" or \
    answer == "n" or answer == "no"):
        answer = input(question).lower().strip()
    if answer[0] == "y":
        return True
    else:
        return False

df = pandas.read_csv(args.csvfile)
count_row = len(df[df['Status'] == False])
skipped_count = len(df[df['Status'] == True])
data = df.to_dict(orient='records')
failedToSend = 0
skippedSendingArray = []
failedToSendArray = []
if yes_or_no("Do you wish to send %d email(s) and skip %d email(s)? (y/n): " % (count_row, skipped_count)):
    pbar = tqdm(total=count_row, file=sys.stdout, unit=" emails")
    for row in data:
        if row['Status'] == True:
            skippedSendingArray.append(row)
            continue
        try:
            send_mail(row)
            pbar.update(1)
        except:
            failedToSend +=1
            failedToSendArray.append(row)
    pbar.close()
    df.to_csv(index=False)
    if skippedSendingArray:
        print("\nSkipped %d row(s) [Hint: You may skip lines with the \"Status\" column with a value of True]" % skipped_count)
        for x in skippedSendingArray:
            print(x)
    if not skippedSendingArray:
        print("\nDid not skip any addresses.")
    
    if failedToSendArray:
        print("\nFailed to send %d email(s) to:" % failedToSend)
        for x in failedToSendArray:
            print(x)
        print("\n[Note: this is only inital sending errors, this does not include the recipient(s) email server rejecting the email or anythig similar.]")
    if not failedToSendArray:
        print("\nSuccessfully sent to all addresses.")

    pbar.close()
else:
    print("Exit")