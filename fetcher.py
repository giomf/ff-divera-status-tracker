from email.message import EmailMessage
from imapclient import IMAPClient
from pathlib import Path
import datetime
import email

DATE_TIME_INPUT_FORMAT = '%a, %d %b %Y %H:%M:%S %z'
DATE_TIME_OUTPUT_FORMAT = '%Y-%m-%dT%H-%M-%S'

def fetch_messages(host: str, user: str, password: str, subject: str, output: Path):
    # context manager ensures the session is cleaned up
    with IMAPClient(host, use_uid=True) as client:
        client.login(user, password)
        client.select_folder('INBOX', readonly=True)

        messages = client.search(['SUBJECT', subject])
        response = client.fetch(messages, ['RFC822'])
        _dettach_files_from_messages(response, output)

def _dettach_files_from_messages(messages, destination: Path):
    for _, data in messages.items():
        message = email.message_from_bytes(data[b'RFC822'], _class=EmailMessage)
        _dettach_file_from_message(message, destination)

def _dettach_file_from_message(message: EmailMessage, destination: Path):
    date = datetime.datetime.strptime(message['Date'], DATE_TIME_INPUT_FORMAT).strftime(DATE_TIME_OUTPUT_FORMAT)
    subject = message['Subject']
    attachment_name = f'{date}_{subject}.xlsx'
    for part in message.walk():
        if part.get('Content-Disposition'):
            attachment_path = destination.joinpath(attachment_name)
            with open(attachment_path, 'wb') as f:
                f.write(part.get_payload(decode=True))
