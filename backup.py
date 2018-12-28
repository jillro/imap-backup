#!/usr/bin/env python3
import argparse
import os
import getpass
from datetime import datetime, timedelta, timezone
from imaplib import IMAP4
import re
import zipfile
from email import parser as emailParser, policy as emailPolicy
from slugify import slugify

parser = argparse.ArgumentParser(description='Backup an IMAP account.')
parser.add_argument('--younger', '--skip-older', type=int, metavar='DAYS', help='Skip messages older than N days.')
parser.add_argument('--older', '--skip-younger', type=int, metavar='DAYS', help='Skip messages younger than N days.')
parser.add_argument('--zip', action='store_true', help='Try to do do incremental archive.')
parser.add_argument('--delete', action='store_true', help='Delete message from server after archiving it.')
args = parser.parse_args()

regex = r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)'
list_response_pattern = re.compile(regex)

hostname = input('Hostname: ')
user = input('Login: ')
password = getpass.getpass()
os.makedirs(os.path.join('output', hostname), exist_ok=True)

if args.zip:
    output = zipfile.ZipFile(os.path.join('output', hostname, user + '-' + datetime.now().strftime("%Y%m%d-%H%M%S") + '.zip'), mode='w')
else:
    output = os.path.join('output', hostname, user + '-' + datetime.now().strftime("%Y%m%d-%H%M%S"))


def parse_list_response(line):
    line = line.decode('UTF-8')
    flags, delimiter, mailbox_name = list_response_pattern.match(line).groups()
    mailbox_name = mailbox_name.strip('"')
    return mailbox_name

def parse_and_save_message(message):
    '''
    Parse the message and write it to the ZipFile

    :param message: The message
    :type message: a dict with headers and body
    '''
    parser = emailParser.BytesFeedParser(policy=emailPolicy.default.clone(refold_source="none", utf8=False))
    parser.feed(message['headers'])
    parser.feed(message['body'])
    email = parser.close()
    subject = email.get('Subject', '')
    try:
        date = email.get('date').datetime
    except AttributeError:
        return False
    if args.younger or args.older:
        if date.tzinfo is not None and date.tzinfo.utcoffset(date) is not None:
            now = datetime.now(timezone.utc)
        else:
            now = datetime.now()
        age_in_days = now - date
        if args.younger and age_in_days > timedelta(days=args.younger):
            return False
        if args.older and age_in_days < timedelta(days=args.older):
            return False
    year = date.strftime('%Y')
    month = date.strftime('%m')
    day = date.strftime('%d')

    filename = date.strftime('%H-%M-%S') + '-' + slugify(subject) + '.eml'
    path = os.path.join(user, box, year, month, day, filename[:255])
    if args.zip:
        output.writestr(path, email.as_bytes())
    else:
        path = os.path.join(output, path)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path,'wb') as f:
            f.write(email.as_bytes())
    return True

with IMAP4(hostname) as host:
    # Connection
    host.starttls()
    host.login(user, password)

    # Get all folders
    boxes = list(map(parse_list_response, host.list()[1]))

    # For each folders
    for box in boxes:
        print('Fetching ' + box)
        rv, data = host.select('"' + box + '"')
        if rv != 'OK':
            print('Unable to open ' + box)
            continue
        count = data[0].decode('UTF-8')
        if count == '0':
            print('Empty folder, skip')
            continue
        print('Downloading ' + count + ' messages')
        # for i in [1, 11, 21, ...]
        for i in range(1, int(count), 10):
            fetch = str(i) + ':' + str(min(i + 9, int(count)))
            print('{:.0%}'.format(i / int(count)), end='\r')
            fetched = host.fetch(fetch, '(BODY[HEADER] BODY[TEXT])')
            if fetched[0] != 'OK':
                raise Exception('Server did not give correct response.')
            messages = fetched[1]
            for j in range(0, len(messages), 3):
                message = {
                    'headers': messages[j][1],
                    'body': messages[j + 1][1]
                }
                if parse_and_save_message(message) and args.delete:
                    message_id = int(i + j/3)
                    host.store(str(message_id), '+FLAGS', '\\Deleted')

    output.close()
    host.expunge()
