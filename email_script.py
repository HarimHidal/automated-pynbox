# -*- coding: utf-8 -*-
"""
Created on Sat Oct 23 19:24:38 2021

@author: harim
"""

######################################################################
# EMAIL HANDLING FUNCTIONS
######################################################################

from smtplib import SMTP
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import imaplib
import email

msg_template = """Hola {name}! Gracias por usar {website}. Estamos muy felices de tenerte con nostros.\n\n""" 
website_URL = "XXXXXXXXX.com"

username = 'XXXXXXXXX@gmail.com'
password = 'XXXXXXXXX'

host = 'imap.gmail.com'

def send(text='Body', subject='Notificación del Equipo XXXXXXXXX', from_email='XXXXXXXXX@gmail.com', to_emails=None, html=None):
    assert isinstance(to_emails, list)
    msg = MIMEMultipart('alternative')
    msg['From'] = from_email
    msg['To'] = ", ".join(to_emails)
    msg['Subject'] = subject
    txt_part = MIMEText(text, 'plain')
    msg.attach(txt_part)
    if html != None:
        html_part = MIMEText(html, 'html')
        msg.attach(html_part)
    msg_str = msg.as_string()
    # login to my smtp server
    server = SMTP(host='smtp.gmail.com', port=587)
    server.ehlo()
    server.starttls()
    server.login(username, password)
    server.sendmail(from_email, to_emails, msg_str)
    server.quit()

def get_inbox(verbose=False):
    mail = imaplib.IMAP4_SSL(host)
    mail.login(username, password)
    mail.select("inbox")
    _, search_data = mail.search(None, 'UNSEEN')
    my_message = []
    for num in search_data[0].split():
        email_data = {}
        _, data = mail.fetch(num, '(RFC822)')
        _, b = data[0]
        email_message = email.message_from_bytes(b)
        for header in ['subject', 'to', 'from', 'date']:
            email_data[header] = email_message[header]
        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True)
                email_data['body'] = body.decode()
            elif part.get_content_type() == "text/html":
                html_body = part.get_payload(decode=True)
                email_data['html_body'] = html_body.decode()
        my_message.append(email_data)
        if verbose:
            for header in ['subject', 'to', 'from', 'date']:
                print("{}: {}".format(header, email_message[header]))
            print("body: " + my_message[0].get("body"))
    return my_message

def format_msg(my_name="", my_website="XXXXXXXXX.com"):
    my_msg = msg_template.format(name=my_name, website=my_website)
    return my_msg 

def send_mail(body, name, website=None, to_email=None, verbose=False):
    if website != None:
        msg = format_msg(my_name=name, my_website=website) + body + "\n\nNOTA: Este es un mensaje automático y no es necesario responder."
    else:
        msg = format_msg(my_name=name) + body + "\n\nNOTA: Este es un mensaje automático y no es necesario responder."
    try:
        send(text=msg, to_emails=[to_email], html=None)
        if verbose:
            print("Email sent to: " + name + " (" + to_email + ")")
        sent = True
    except:
        if verbose:
            print("Error while sending email to: " + name + " (" + to_email + ")")
        sent = False
    return sent

######################################################################
# SPREADSHEETS HANDLING FUNCTIONS
######################################################################

from googleapiclient.discovery import build
from google.oauth2 import service_account

spreadsheet_ID = 'XXXXXXXXXXXXXXXXXX'

def read_spreadsheet(RANGE_TO_READ, SPREADSHEET_ID, VERBOSE=False):
    
    SERVICE_ACCOUNT_FILE = 'key.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    creds = None
    creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    service = build('sheets', 'v4', credentials=creds)
    
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_TO_READ).execute()
    values = result.get('values', [])
    if VERBOSE:
        print(values)
    return values

def write_spreadsheet(LIST_TO_WRITE, RANGE_TO_WRITE, SPREADSHEET_ID, VERBOSE=False):
    
    SERVICE_ACCOUNT_FILE = 'key.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    creds = None
    creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    
    service = build('sheets', 'v4', credentials=creds)
    valores = LIST_TO_WRITE
    
    # Call the Sheets API
    sheet = service.spreadsheets()
    request = sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=RANGE_TO_WRITE, valueInputOption="USER_ENTERED", body={"values":valores}).execute()
    if VERBOSE:
        print(request)
    
    
######################################################################
# CALLS OF THE FUNCTIONS (MAIN)
######################################################################

# ANSWER THE MAILS RECEIVED (INBOX) WITH NO-REPLY MAILS

no_reply_msg = "Te agradecemos por intentar contactarnos, lamentablemente, la bandeja de entrada de este correo no está monitoreada. Si tienes alguna duda por favor visita nuestra página web."

try:
    print("***** PROCESS: SEND NO-REPLY EMAILS *****")
    raw_data = get_inbox()
    inbox = []
    nr_mailing_list = []
    for email_received in raw_data:
        inbox.append(email_received)
    for mail in inbox:
        nr_mailing_list.append(mail['from'])
    for mail in nr_mailing_list:
        try:
            send_mail(no_reply_msg, "de nuevo", website_URL, mail)
        except:
            print("ERROR: unable to send no-reply email to " + mail[1] + ". Set verbose=True to perform a BTS.")
    print("*****************************************")
except:
    print("ERROR: unable to retrieve the data from the inbox. Set verbose=True to perform a BTS.\n*****************************************")

# CONFIRM USERS THE INFO WAS CORRECTLY SUBMITTED TO GOOGLE FORMS

try:    
    print("**** PROCESS: SEND CONFIRMATION MAIL ****")
    
    range_to_read = "Form Responses 1!A1:L1000"
    db = read_spreadsheet(range_to_read, spreadsheet_ID)
    db.pop(0)
    gf_msg = "Ya recibimos tu solicitud para XXXXXXXXX en {}"
    mailing_list = []
    control = []
    i = 1
    for entry in db:
        i=i+1
        control.append(i)
        if int(entry[11])==0:
              mailing_list.append([entry[1], entry[4], entry[7]])
        else:
              control.pop(control.index(i))
    if mailing_list != []:
        for mail in mailing_list:
            try:
                send_mail(gf_msg.format(mail[1]), mail[0], website_URL, mail[2])
                write_spreadsheet([["1"]], "L"+ str(control.pop(0)), spreadsheet_ID)
            except:
                print("ERROR: unable to send confirmation email to " + mailing_list[2] + ". Set verbose=True to perform a BTS.")
    print("*****************************************")
except:
    print("ERROR: unable to send the Google Forms confirmation emails. Set verbose=True to perform a BTS.\n*****************************************")     
    





