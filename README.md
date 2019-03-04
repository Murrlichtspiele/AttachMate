# AutomateKDM
Downloading and unpacking mail attachments in a mailbox subfolder
# Use case
This script is intended to download KDMs (key delivery messages) for cinema servers from a mail server, unpack them (they are usually packaged in zip files) and ingest them into the cinema server via FTP or SSH. This script is under development at the moment!
## TO-DO:
- :white_check_mark: Write a attachment parsers for IMAP servers that downloads new KDMs (
	- :white_check_mark: download works. Uses python imaplib.
	- :white_check_mark: Distinction between old and new is yet to be implemented.

- :white_check_mark: Unpack the ZIP files into a work directory.
- :white_check_mark: Find out how to automate the ingest process.
  - :white_check_mark: Server can be started via WOL. FTP can put KDMs on the local drive. Ingest via SSH possible?
  - How to start the projector, aka wake him from sleep? (IMB needed for KDM ingest to work)
- Send a conformation mail with the newly ingested KDMs (EXTRA: Add the expiration dates in the mail.)  

