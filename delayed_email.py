#!/usr/local/bin/env python3
"""
Script that sends an email after or on a specific time.

todo add more exception handling
todo Optimize
todo simplify
"""
from smtplib import SMTP
from imaplib import IMAP4_SSL
import email
import os

imap_server = 'imap-mail.outlook.com'
imapport = 993
smtp_server = 'smtp-mail.outlook.com'
smtpport = 587
user = os.environ['OUTLOOK_USER']
pwd = os.environ['OUTLOOK_PASS']
folder = 'delayed'


def send_delayed_msg():
    """
    """
    with IMAP4_SSL(imap_server, imapport) as i:
        i.login(user, pwd)
        i.select(folder, readonly=False)
        status, msg_ids = i.search(None, 'ALL')
        if status == 'OK':
            if msg_ids != [b'']:
                for id in msg_ids[0].split():
                    status, data = i.fetch(id, '(RFC822)')
                    if status == 'OK':
                        raw_msg = data[0][1]
                        email_msg = email.message_from_bytes(raw_msg)
                    else:
                        print('Status:', status, data)
                    # change so that all msg sent in one connection to smtp?
                    with SMTP(smtp_server, smtpport) as s:
                        s.starttls()
                        s.login(user, pwd)
                        # add try? to catch exception. does with do the same?
                        s.send_message(email_msg)

                    typ, data = i.store(id, '+FLAGS', '\\Deleted')
                    if typ != 'OK':
                        print('Status:', typ)
            else:
                print('Folder is empty. No messages to send.')
        else:
            print('Status:', status,
                  'Something went wrong while retrieving messages ids.')
        i.expunge()


send_delayed_msg()

# todo send me a confirmation, log that 1 the script started and there was no error during the sending so I know everything went according to plan.
"""Send confirmation report
"""
