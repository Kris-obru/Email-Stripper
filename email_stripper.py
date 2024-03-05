import logging
import shutil
import json
import os
from pathlib import Path
from email import message_from_file
from email.iterators import typed_subpart_iterator
from email.message import Message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import parsedate_to_datetime

# Source and destination directories

filename = "./config.json"

default_settings = {
    "source_dir" : "./emails",
    "destination_dir" : "./fixed",
    "attach_dir" : "./attachments",
    "size_in_bytes" : 25165824
}    

try:
    with open(filename, "r") as file:
        data = json.load(file)
except FileNotFoundError:
    with open(filename, "w") as file:
        json.dump(default_settings, file, indent=4)
    data = default_settings

source_dir = data["source_dir"]
destination_dir = data["destination_dir"]
attach_dir = data["attach_dir"]
size_in_bytes = data["size_in_bytes"]

def sanitize(name):
    # Replace invalid characters in the folder name
    invalid_chars = set('\\/:*?"<>|\n')
    return ''.join(char if char not in invalid_chars else '_' for char in name)

def join(a,b): # easier to do this
    return os.path.join(a,b)

def makedir(path): # easy way to create dir without doing it in the main function
    if not os.path.exists(path):
        os.makedirs(path)

def copy(src,dest): # easy way to copy dir without doing it in the main function
    if not os.path.exists(dest):
        shutil.copytree(src, dest)

def extract_email_from_path(path):
    # Normalize the path to remove redundant separators and up-level references
    normalized_path = os.path.normpath(path)

    # Split the path into its components
    parts = normalized_path.split(os.sep)

    # Find the first part that looks like an email address
    for part in parts:
        if '@' in part:
            return part

    return None

def extract_attachments(eml_file_path, stripped_eml_dir, attachments_dir):
    # Parse the .eml file
    with open(eml_file_path, 'r', encoding='utf-8-sig') as eml_file:
        msg = message_from_file(eml_file)

    # Extract the date from the Date header
    date_header = msg.get('Date')
    if date_header:
        date = parsedate_to_datetime(date_header)
        if date:
            date_string = date.strftime('%Y-%m-%d_%H-%M-%S')
        else:
            date_string = 'unknown'
    else:
        date_string = 'unknown'

    # Create a new message without attachments
    new_msg = MIMEMultipart()
    new_msg['From'] = msg['From']
    new_msg['To'] = msg['To']
    new_msg['Subject'] = msg['Subject']

    # Iterate through message parts
    for part in typed_subpart_iterator(msg, 'multipart', 'alternative'):
        for subpart in part.get_payload():
            if isinstance(subpart, Message):
                new_msg.attach(subpart)

    # Save the new message without attachments to a new .eml file
    stripped_eml_file_path = os.path.join(stripped_eml_dir, os.path.basename(eml_file_path))
    with open(stripped_eml_file_path, 'w') as stripped_eml_file:
        stripped_eml_file.write(new_msg.as_string())

    # Move attachments to the attachments directory
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        filename = part.get_filename()
        if filename:
            # Remove invalid characters from the filename
            filename = sanitize(filename)
        
            makedir(join(attachments_dir, date_string))
            with open(os.path.join(attachments_dir, date_string, filename), 'wb') as attachment_file:
                if part.get_payload(decode=True) is not None:
                    attachment_file.write(part.get_payload(decode=True))

def process_directory(directory):
    directory = os.path.normpath(directory)  # Normalize the directory path
    for root, dirs, files in os.walk(directory):
        logging.info("[PROCESSING]: " + root)
        for file in files:
            if file.endswith('.eml'):
                # Process the .eml file here
                file_path = os.path.join(root, file)
                if os.path.getsize(file_path) >= size_in_bytes:
                    fixed_dir = os.path.dirname(os.path.join(destination_dir, os.path.relpath(file_path, source_dir)))
                    email = extract_email_from_path(fixed_dir)
                    extract_attachments(file_path, fixed_dir, os.path.join(attach_dir, email))

# Configure logging
logging.basicConfig(filename='email_processing.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
console = logging.StreamHandler()
console.setLevel(logging.INFO)
logging.getLogger('').addHandler(console)


if not os.path.exists(source_dir):
    makedir(source_dir)
if not os.path.exists(destination_dir):
    makedir(destination_dir)
if not os.path.exists(attach_dir):
    makedir(attach_dir)


for sender_folder_name in os.listdir(source_dir): # loops through each email in ./emails
    sender_folder_path = join(source_dir, sender_folder_name) # path for email in ./emails
    sender_fixed_path = join(destination_dir, sender_folder_name) # path for email in ./fixed
    sender_attach_path = join(attach_dir, sender_folder_name)
    
    logging.info("[INFO]: Copying " + sender_folder_name)
    copy(sender_folder_path, sender_fixed_path) # copies email from ./emails to ./fixed

    logging.info("[INFO]: Creating attachments path for " + sender_folder_name)
    makedir(sender_attach_path)

    logging.info("[INFO]: Processing " + sender_folder_name)
    process_directory(sender_folder_path)

logging.info("[YEAH BABY]: Finished!")
