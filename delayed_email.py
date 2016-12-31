#!/usr/local/bin/env python3
"""
Script that sends email messages that were saved to a mailbox folder.
When a specific condition (time or other) is reached the script is launched.

Used with launchd daemon on OSX.
Tested with outlook.com.
Username and password stored in environment variables.
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

# info for confirmation email
send_from = user
send_confirm_to = user
msg_body = []
success = False


def get_msg_list(msg_id_list, imap_connection):
    """Gets a list of imap messages id and an active imap connection
    then return a list of email.message.message"""
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
    """Gets a list of email.message.message and send each of them"""
    with SMTP(smtp_server, smtpport) as s:
        s.starttls()
        s.login(user, pwd)
        for msg in msg_list:
            s.send_message(msg)
            msg_body.append(
                msg.get('subject') + ' was sent to ' + msg.get('to'))


def delete_msgs(msg_id_list, imap_connection):
    """Gets a list of imap messages id and an active imap connection and
    deleted those message. This happens in the imap connection selected
    folder."""
    for id in msg_id_list:
        typ, data = imap_connection.store(id, '+FLAGS', '\\Deleted')
        if typ != 'OK':
            print('Status:', typ, 'Problem deleting messages.')


def send_confirmation():
    """Send an email confirmation message"""
    msg_subject = "Subject: Your delayed emails were sent"
    msg = msg_subject + '\n'
    for item in msg_body:
        msg += item + '\n'
    with SMTP(smtp_server, smtpport) as s:
        s.starttls()
        s.login(user, pwd)
        s.sendmail(send_from, send_confirm_to, msg)


with IMAP4_SSL(imap_server, imapport) as i:
    i.login(user, pwd)
    i.select(folder, readonly=False)

    status, msg_ids = i.search(None, 'ALL')
    if status == 'OK' and msg_ids != [b'']:
        msg_id_list = msg_ids[0].split()
        msg_list = get_msg_list(msg_id_list, i)
        if send_email_msgs(msg_list) is None:
            if delete_msgs(msg_id_list, i) is None:
                i.expunge()
                success = True
            else:
                print('There was a problem deleting the messages.')
        else:
            print('There was a problem sending the messages.')
    else:
        if status == 'OK':
            print('Folder is empty. No messages to send.')
        else:
            print('Something went wrong retrieving messages list.')

if success is True:
    send_confirmation()
else:
    pass

# TODO: Save to log.
# TODO: better exception handling
# TODO: multiple folders
# TODO: unit tests
# TODO: Optimize
# TODO: simplify
