from simplegmail import Gmail
from simplegmail.query import construct_query
import re
from weasyprint import HTML
import os
import sys
import tkinter as tk
from tkinter import filedialog
import json
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
import pickle
from google.auth.transport.requests import Request

# Function to select a folder and save its path to the config.json file
def choose_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    selected_folder = filedialog.askdirectory(title="Select Folder to Save Invoices")  # Prompt user to select a folder
    if selected_folder:
        # Save the selected folder path to the config.json file
        with open(config_file_path, 'w') as config_file:
            json.dump({"invoices_folder": selected_folder}, config_file)
        print(f"Selected folder: {selected_folder}")
        return selected_folder
    else:
        print("No folder selected.")
        sys.exit()

# Function to authenticate and get Gmail credentials
def get_gmail_credentials():
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
    credentials_file = 'client_secret.json'  # Path to your OAuth 2.0 credentials file
    token_file = 'token.pickle'  # File where access tokens will be stored

    creds = None
    if os.path.exists(token_file):
        with open(token_file, 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_file, 'wb') as token:
            pickle.dump(creds, token)

    return creds

# Determine the base path
if getattr(sys, 'frozen', False):
    # Running as a PyInstaller bundle
    base_path = sys._MEIPASS
else:
    # Running as a normal Python script
    base_path = os.path.dirname(os.path.abspath(__file__))

# Path to the JSON file where folder path will be saved
config_file_path = os.path.join(os.path.expanduser("~"), 'config.json')

# Load the folder path from config.json file if it exists
if os.path.exists(config_file_path):
    with open(config_file_path, 'r') as config_file:
        config_data = json.load(config_file)
        invoices_folder = config_data.get("invoices_folder", None)
        if not invoices_folder or not os.path.exists(invoices_folder):
            invoices_folder = choose_folder()
else:
    # If config.json doesn't exist, prompt the user to select a folder
    invoices_folder = choose_folder()

# Get Gmail credentials
creds = get_gmail_credentials()
gmail = Gmail()  # Pass the credentials to the Gmail class

query_params = { # select the list of email you want to get from the Gmail inbox
    "newer_than": (7, "day"),
    #"unread": False,
}

mails = gmail.get_messages(query=construct_query(query_params)) # run the query and get list of emails

for message in mails:
    safe_filename = re.sub(r'[<>:"/\\|?*]', '_', message.subject + message.date) # change the file name to a safe name so you can save it on your pc
    file_path = invoices_folder + '\\' + safe_filename + '.pdf' # set the new pdf name
    if os.path.exists(file_path): # checks if the pdf name already exists in the directory, if not continue
        print('All invoices are saved!')
        break
    else: 
        if any(word in message.subject for word in ['חשבונית', 'invoice', 'receipt', 'קבלה', 'bill']): # if the email subject contains these words continue 
            if message.attachments: # if the mail contains pdf/image/file save the attachments
                for attm in message.attachments:
                    attm.save(filepath=file_path) # save the file
                    print('Saved attachment of : ' + safe_filename)
        
            elif message.html: # if the email doesn't contain attachments, save the mail content as pdf
                html_content = message.html
                pdf_file_name = message.subject + message.date # Specify the name of the PDF file

                try:
                    HTML(string=html_content).write_pdf(file_path, optimize_size=False)# Convert the HTML content to PDF and save it
                    print('Saved: ' + safe_filename + " content as pdf")
                except Exception:
                    print('Cannot save this mail as html: ' + safe_filename)
                else:
                    continue
        else:
            continue
