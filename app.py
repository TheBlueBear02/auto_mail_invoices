from simplegmail import Gmail
from simplegmail.query import construct_query
import re
from weasyprint import HTML
import os


gmail = Gmail() # will open a browser window to ask you to log in and authenticate for the first time
save_directory = r'D:\Projects\emailAPI\invoices' # the saved invoices directory

query_params = { # select the list of email you want to get from the Gmail inbox
    "newer_than": (2, "day"),
    #"unread": False,
}

messages = gmail.get_messages(query=construct_query(query_params)) # run the query and get list of emails

for message in messages:
    safe_filename = re.sub(r'[<>:"/\\|?*]', '_', message.subject) # change the file name to a safe name so you cna save it on your pc
    file_path = save_directory + '\\' + safe_filename + '.pdf' # set the new pdf name
    if os.path.exists(file_path): # checks if the pdf name already exists in the directory, if not continue
        print('All invoices are saved!')
        break
    else: 
        if any (word in message.subject for word in ['חשבונית', 'invoice','receipt','קבלה','bill']): # if the email subject contains these words continue 
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