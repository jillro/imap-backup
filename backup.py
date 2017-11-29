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
args = parser.parse_args()

regex = r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)'
list_response_pattern = re.compile(regex)

hostname = input('Hostname: ')
user = input('Login: ')
password = getpass.getpass()
os.makedirs(os.path.join('output', hostname), exist_ok=True)
output = zipfile.ZipFile(os.path.join('output', hostname, user + datetime.now().strftime("%Y%m%d-%H%M%S") + '.zip'), mode='w')


def parse_list_response(line):
    line = line.decode('UTF-8')
    flags, delimiter, mailbox_name = list_response_pattern.match(line).groups()
    mailbox_name = mailbox_name.strip('"')
    return mailbox_name


def parse_fetch_response(box, messages):
    for j in range(0, len(messages), 3):
        parser = emailParser.BytesFeedParser(policy=emailPolicy.default.clone(refold_source="none", utf8=False))
        parser.feed(messages[j][1])
        parser.feed(messages[j + 1][1])
        email = parser.close()
        subject = email.get('Subject', '')
        try:
            date = email.get('date').datetime
            age_in_days = datetime.now(timezone.utc) - date
            if args.younger and age_in_days > timedelta(days=args.younger):
                continue
            if args.older and age_in_days < timedelta(days=args.older):
                continue
        except AttributeError:
            continue
        year = date.strftime('%Y')
        month = date.strftime('%m')
        day = date.strftime('%d')

        filename = date.strftime('%H-%M-%S') + '-' + slugify(subject) + '.eml'
        output.writestr(os.path.join(user, box, year, month, day, filename), email.as_bytes())

with IMAP4(hostname) as host:
    # Connection
    host.starttls()
    host.login(user, password)
    # Get all folders
    boxes = list(map(parse_list_response, host.list()[1]))

    # For each folders
    for box in boxes:
        print('Fetching ' + box)
        count = host.select('"' + box + '"')[1][0].decode('UTF-8')
        if count == '0':
            print('Empty folder, skip')
            continue
        print('Downloading ' + count + ' messages')
        # for i in [1, 11, 21, ...]
        for i in range(1, int(count), 10):
            fetch = str(i) + ':' + str(min(i + 9, int(count)))
            print('{:.0%}'.format(i / int(count)), end='\r')
            messages = host.fetch(fetch, '(BODY[HEADER] BODY[TEXT])')[1]
            parse_fetch_response(box, messages)
