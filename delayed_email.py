#!/usr/local/bin/env python3
"""
Script that sends saved email after or on a specific time.

Tested with outlook.com.
username and password stored in environment variables.
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


def get_msg_list(msg_id_list, imap_connection):
    msg_list = []
    for id in msg_id_list:
        status, data = imap_connection.fetch(id, '(RFC822)')
        if status == 'OK':
            raw_msg = data[0][1]
            email_msg = email.message_from_bytes(raw_msg)
            msg_list.append(email_msg)
        else:
            print('Status:', status, 'Error retrieving raw message')
    return msg_list


def send_email_msgs(msg_list):
    if len(msg_list) > 0:
        with SMTP(smtp_server, smtpport) as s:
            s.starttls()
            s.login(user, pwd)
            for msg in msg_list:
                s.send_message(msg)
    else:
        print('List is empty.')


def delete_msgs(msg_id_list, imap_connection):
    for id in msg_id_list:
        typ, data = imap_connection.store(id, '+FLAGS', '\\Deleted')
        if typ != 'OK':
            print('Status:', typ, 'Problem deleting messages.')


with IMAP4_SSL(imap_server, imapport) as i:
    i.login(user, pwd)
    i.select(folder, readonly=False)
    status, msg_ids = i.search(None, 'ALL')
    if status == 'OK':
        if msg_ids != [b'']:
            msg_id_list = msg_ids[0].split()
            msg_list = get_msg_list(msg_id_list, i)
            if len(msg_list) > 0:
                if send_email_msgs(msg_list) is None:
                    if delete_msgs(msg_id_list, i) is None:
                        i.expunge()
                    else:
                        print('There was a problem deleting the messages.')
                else:
                    print('There was a problem sending the messages.')
            else:
                print(
                    'Something went wrong. List is empty but folder is not.')
        else:
            print('Folder is empty. No messages to send.')
    else:
        print('Status:', status,
              'Something went wrong while retrieving messages ids from',
              folder)

# todo send me a confirmation, log that 1 the script started and there was no error during the sending so I know everything went according to plan.
"""Send confirmation report
"""

# todo add more/better exception handling
# todo unit tests
# todo Optimize
# todo simplify
