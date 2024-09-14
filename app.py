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
import pdfplumber
import openai
from datetime import datetime
from googleapiclient.discovery import build
from google.oauth2 import service_account

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

def extract_text_from_pdf(pdf_path): # get pdf file and return the file's text
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text()
        return full_text

def extract_invoice_details(invoice_text):
    """
    Function to extract company name and amount paid from the invoice text using OpenAI Chat API.
    """
    messages = [
        {"role": "system", "content": "You are an assistant that extracts amount paid from each invoice the user paid."},
        {"role": "system", "content": f"תוציא את סך הכל (סה''כ) הסכום ששולם בחשבונית."},
        {"role": "user", "content": f"Extract only the bottom line of what paid from this invoice. The text may be in English or Hebrew (return only the amouint paid in numbers or if you not sure return None):\n\n{invoice_text}"}
    ]

    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",  # You can also use "gpt-4" for more accurate results
            messages=messages,
            max_tokens=150,
            temperature=0.1,
        )
        
        return response['choices'][0]['message']['content'].strip()

    except Exception as e:
        print(f"Error during OpenAI API call: {str(e)}")
        return None

def save_in_sheets(values):
    # Define the scopes required
    scopes = ['https://www.googleapis.com/auth/spreadsheets']

    # Authenticate and create the service object
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=scopes)
    service = build('sheets', 'v4', credentials=creds)

    # Define the ID of the Google Sheet and the range you want to update
    SPREADSHEET_ID = '1iX-UxZhkscNm3pkDT-9m-RKB0dOKhiW_qiyqZ9ZO0wE'
    RANGE_NAME = '2024!A2'

    body = {
        'values': values
    }

    # Write data to the sheet
    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=RANGE_NAME,
        valueInputOption='RAW',
        insertDataOption='INSERT_ROWS',
        body=body
    ).execute()     

# Get Gmail credentials
creds = get_gmail_credentials()
gmail = Gmail()  # Pass the credentials to the Gmail class
# Get openai api key
with open(r'D:\Projects\auto_mail_invoices\openai_key.json', 'r') as openai_file:
        openai.api_key = json.load(openai_file)['key']

# Load your credentials from the downloaded JSON file
SERVICE_ACCOUNT_FILE = r'D:\Projects\auto_mail_invoices\service_account.json'


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



query_params = { # select the list of email you want to get from the Gmail inbox
    "newer_than": (31, "day"),
    #"unread": False,
}

mails = gmail.get_messages(query=construct_query(query_params)) # run the query and get list of emails

for message in mails:
    # Changing the date format to be more clear
    date_string = message.date.split()[0]
    date_obj = datetime.strptime(date_string, '%Y-%m-%d')
    clear_date = date_obj.strftime('%d-%m-%Y')
    clear_time = ":".join(message.date.split()[1].split(":")[:2])
    clear_time_date = re.sub(r'[<>:"/\\|?*]', '-', clear_date + '_' + clear_time)
    safe_filename = re.sub(r'[<>:"/\\|?*]', '_', message.subject + '_' + clear_time_date) # change the file name to a safe name so you can save it on your pc
    
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
                    invoice_text = extract_text_from_pdf(file_path) # Gets the text of the pdf
                    amount = extract_invoice_details(invoice_text) # Extract the details from the invoice
                    
                    # Data to append to Google Sheets
                    data_to_append = [[message.sender,'אימייל', amount,'חדש', clear_date, file_path]]
                    save_in_sheets(data_to_append)
                    print('Saved in Google Sheets!')
        
            elif message.html: # if the email doesn't contain attachments, save the mail content as pdf
                html_content = message.html
                try:
                    HTML(string=html_content).write_pdf(file_path, optimize_size=False)# Convert the HTML content to PDF and save it
                    print('Saved: ' + safe_filename + " content as pdf")
                    invoice_text = extract_text_from_pdf(file_path)
                    amount = extract_invoice_details(invoice_text) # Extract the details from the invoice
                    print(message.sender)
                    print(amount)

                except Exception:
                    print('Cannot save this mail as html: ' + safe_filename)
                else:
                    continue
        else:
            continue

input("Click Enter to exit")