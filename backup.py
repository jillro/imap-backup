#!/usr/bin/env python3
import argparse
import json
import os
import getpass
from datetime import datetime, timedelta, timezone
import re
import zipfile
from email import parser as emailParser, policy as emailPolicy

from imapclient import IMAPClient
from imapclient.exceptions import IMAPClientError
from slugify import slugify

parser = argparse.ArgumentParser(description='Backup an IMAP account.')
parser.add_argument('--younger', '--skip-older', type=int, metavar='DAYS', help='Skip messages older than N days.')
parser.add_argument('--older', '--skip-younger', type=int, metavar='DAYS', help='Skip messages younger than N days.')
parser.add_argument('--zip', action='store_true', help='Try to do do incremental archive.')
parser.add_argument('--delete', action='store_true', help='Delete message from server after archiving it.')
parser.add_argument('--retry', default=None, help="Specify a retry file")
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

def parse_and_save_message(message):
    '''
    Parse the message and write it to the ZipFile

    :param message: The message
    :type message: a dict with headers and body
    '''
    parser = emailParser.BytesFeedParser(policy=emailPolicy.default.clone(refold_source="none", utf8=False))
    parser.feed(message[b'BODY[HEADER]'])
    parser.feed(message[b'BODY[TEXT]'])
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
    path = os.path.join(user, box_name, filename[:255])
    if args.zip:
        output.writestr(path, email.as_bytes())
    else:
        path = os.path.join(output, path)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path,'wb') as f:
            f.write(email.as_bytes())
    return True

with IMAPClient(hostname) as host:
    # Connection
    host.login(user, password)

    # Get all folders
    boxes = host.list_folders()

    if args.retry is not None:
        completed_boxes = json.loads(open(args.retry))["completed_boxes"]
    else:
        completed_boxes = []
    # For each folders
    for flags, delimiter, box_name in boxes:
        print('Fetching ' + box_name)
        try :
            count = host.select_folder(box_name)[b'EXISTS']
            print('Downloading ' + str(count) + ' messages')
            messages = host.search()
            for i in range(0, int(count), 10):
                messages_slice = messages[i:min(i+10, count)]
                print('{:.0%}'.format(i / count), end='\r')
                fetched = host.fetch(messages_slice, ['BODY[HEADER] BODY[TEXT]'])
                for id, message in fetched.items():
                    if parse_and_save_message(message) and args.delete:
                        message_id = int(i + id/3)
                        host.store(str(message_id), '+FLAGS', '\\Deleted')

            completed_boxes.append(box_name)
        except IMAPClientError:
            with open(user + '-' + datetime.now().isoformat() + '-retry.json', 'w') as fp:
                json.dump({'completed_boxes': completed_boxes}, fp)


    host.expunge()
