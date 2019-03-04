#!/usr/bin/python
# Use case: Download KDMs from mail account and manage the ingest of them.
# Needed:
# - Download mail payload to folder.
# - Unpack ZIP files.
# - Wake up Doremi and projector and start the ingest
# - Make Doremi check for new movies to ingest
# - Send info mail with ingested kdms and started DCP ingests.

import email
import getpass, imaplib
import os, zipfile
import sys, re
from ftplib import FTP
import platform   # For getting the operating system name
import subprocess  # For executing a shell command
import pyping
from wakeonlan import send_magic_packet
import xml.etree.ElementTree as ET
import time 

imapServer = 'imap.some.server'
userName = 'example@some.server'
passwd = getpass.getpass('Enter password for ' + userName + ' at ' + imapServer + ': ')
subfolder = 'Inbox'
detach_dir ='/tmp/'

pattern_uid = re.compile('\d+ \(UID (?P<uid>\d+)\)')

def connect(pw):
    imapSession = imaplib.IMAP4_SSL(imapServer)
    result, accountDetails = imapSession.login(userName, pw)
    if result != 'OK':
        print 'Not able to sign in!'
        raise
    else:
        print 'Sucessfully logged in!'
        return imapSession

def parse_uid(data):
    match = pattern_uid.match(data)
    return match.group('uid')

#####################################################################
#  Check if mail server and doremi are reachable before continuing  #
#####################################################################

r = pyping.ping('192.168.1.129')

if r.ret_code == 0:
    print 'Doremi is up, continuing...'
    up = True
else:
    print 'Doremi does not respond to pin request. Sending WOL signal'
    up = False
    send_magic_packet('00.25.90.7E.5B.6C')
    time.sleep(180)
    r = pyping.ping('192.168.1.129')
    if r.ret_code == 0:
        print 'Doremi is up, continuing...'
    else:
        print 'Doremi is still not responding. Something went wrong. (no power?)'
    #WOL
    
# Create the output directory:
if 'attachments' not in os.listdir(detach_dir):
    os.mkdir(detach_dir + 'attachments')

# Connect to server
imapSession = connect(passwd)

# Dive into inbox
imapSession.select(subfolder)

# Get all mails in the inbox
result, mails = imapSession.search(None, 'ALL')
if result != 'OK':
    print 'Error searching Inbox.'
    raise

# Iterating over all mails
forlater = mails[0].split()
for msgId in mails[0].split():
    print 'Working on msgId ('+msgId+')'
    result, messageParts = imapSession.fetch(msgId, '(RFC822)')
    
    if result != 'OK':
        print 'Error fetching mail.'
    
    emailBody = messageParts[0][1]
    mail = email.message_from_string(emailBody)
    
    for part in mail.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        
        fileName = part.get_filename()
        
        # Only download .zip or .xml files
        if fileName.endswith('.zip'):
            print 'Found ZIP file ' + fileName
        elif fileName.endswith('.xml'):
            print 'Found XML file ' + fileName
        else:
            print 'Omitting file ' + fileName
            continue
        
        if bool(fileName):
            filePath = os.path.join(detach_dir, 'attachments', fileName)
            if not os.path.isfile(filePath):
                print fileName
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
uids = []
for msgId in forlater:
    result, data = imapSession.fetch(msgId, "(UID)")
    msg_uid = parse_uid(data[0])
    uids.append(msg_uid)
for uid in uids:
    result = imapSession.uid('COPY', uid,'INBOX.old')
    print result
    if result[0] == 'OK':
        mov, data = imapSession.uid('STORE', uid , '+FLAGS', '(\Deleted)')
        imapSession.expunge()
        print 'Deleted'
    time.sleep(2)

imapSession.close()
imapSession.logout()

for filename in os.listdir(detach_dir + 'attachments'):
    if filename.endswith('.zip'):
        file_path = detach_dir + 'attachments/' + filename
        print "Extracting " + filename
        zip = zipfile.ZipFile(file_path)
        try:
            zip.extractall(path=detach_dir + 'attachments')
        except:
            print 'Could not extract zip file ' + attachments
        try:
            os.remove(file_path)
        except:
            print 'Could not remove ' + file_path
ingested_kdms = []

if os.listdir(detach_dir + 'attachments') == []:
    print('No KDMs downloaded.')
else:
    print ('I downloaded the following KDMs:')
    for nr, filename in enumerate(os.listdir(detach_dir + 'attachments')):
        print(str(nr) + ': '+ str(filename))
    
    ftp = FTP('cranky')
    ftp.login('ingest','password') # Not the real password 
    serverdirectorypath='/'
    serverfilepath=serverdirectorypath

    for filename in os.listdir(detach_dir + 'attachments'):
        file_path = detach_dir + 'attachments/' + filename
        try:
            ftp.storbinary('STOR %s'%serverfilepath + filename, open(file_path, 'rb'))
            ingested_kdms.append(filename)
        except:
            print('Could not ingest KDM' + filename)
        os.remove(file_path)

print(ingested_kdms)
