# Use case: Download KDMs from mail account and manage the ingest of them.
# Needed: 
# - Download mail payload to folder.
# - Unpack ZIP files.
# - Wake up Doremi and projector and start the ingest
# - Make Doremi check for new movies to ingest
# - Send info mail with ingested kdms and started DCP ingests.

import email
import getpass, imaplib
import os
import sys

# Create the output directory:
detach_dir = '.'
if 'attachments' not in os.listdir(detach_dir):
    os.mkdir('attachments')

imapServer = 'imap.strato.de'
userName = 'kdm@murrlichtspiele.de'
passwd = getpass.getpass('Enter password for ' + userName + ' at ' + imapServer + ': ')

subfolder = 'Inbox'

try:
    imapSession = imaplib.IMAP4_SSL(imapServer)
    typ, accountDetails = imapSession.login(userName, passwd)
    if typ != 'OK':
        print 'Not able to sign in!'
        raise
    else:
        print 'Sucessfully logged in!'
    
except:
    print 'Not able to connect to server ' + imapServer + ' with username '+ userName

    
try:    
    imapSession.select(subfolder)
    typ, data = imapSession.search(None, 'ALL')
    if typ != 'OK':
        print 'Error searching Inbox.'
        raise
except :
    print 'Not able to dive into subfolder.'

try:    
    # Iterating over all emails
    for msgId in data[0].split():
        typ, messageParts = imapSession.fetch(msgId, '(RFC822)')
        if typ != 'OK':
            print 'Error fetching mail.'
            raise

        emailBody = messageParts[0][1]
        mail = email.message_from_string(emailBody)
        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                # print part.as_string()
                continue
            if part.get('Content-Disposition') is None:
                # print part.as_string()
                continue
            fileName = part.get_filename()
            print fileName
            if fileName.endswith('.zip'):
                print 'Found ZIP file. Unpacking will happen later.'
            elif fileName.endswith('.xml'):
                print 'Found XML file.'
            else:
                print 'Omitting filetype.'
                continue
            if bool(fileName):
                filePath = os.path.join(detach_dir, 'attachments', fileName)
                if not os.path.isfile(filePath) :
                    print fileName
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
    imapSession.close()
    imapSession.logout()
except :
    print 'Not able to download all attachments.'
