
import base64
import os.path
import pprint
import openpyxl

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.text import MIMEText

from email.message import EmailMessage

import information

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def gmail_send_message(service,
                       to_mail:list,
                       from_mail:str,#
                       message_theme:str,
                       message_text:str,
                        html = None):
    if not service:
        return "Нет сервиса."

    try:
        service = service
  
        for email in to_mail: 
            message = EmailMessage()
            message.set_content(message_text)

            if html:
                message = MIMEText(html,"html")
        
            message['Subject'] = message_theme
            message['From'] = from_mail
            message['To'] = email
                
         
            encoded_message = base64.urlsafe_b64encode(message.as_bytes()) \
                        .decode()

            create_message = {
                        'raw': encoded_message
                    }
            try:
                send_message = (service.users().messages().send(userId="me", body=create_message).execute())
                print(f"to: {email}, status: Sucсessfuly")   

            except:
                print(f"to: {email}, status: Error")
            
    except HttpError as error:
        print(F'An error occurred: {error}')
        send_message = None
        
        

def return_service():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
   
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'cred.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('gmail', 'v1', credentials=creds)
        return service
    
    except HttpError as error:
        print(f'An error occurred: {error}')
        return None



def parse_excel_file(path:str,sheet_name:str,column:int, )->list:
    workbook = openpyxl.load_workbook(path)
    sheet = workbook[sheet_name]
    column_number = column
    column_data = []
   
    for row in sheet.iter_rows(min_col=column_number, max_col=column_number, values_only=True):
        for cell_value in row:
            column_data.append(cell_value)
            # print(type(cell_value))
    
    workbook.close()    
    return column_data

def main():

    service = return_service()
   
    test = parse_excel_file('emails.xlsx',sheet_name="тест",column=1)
    

    gmail_send_message( service,
                        to_mail = test,
                        from_mail="jorj.knyazyan.15@gmail.com",
                        message_text=information.text, 
                        message_theme=information.theme,
                        )
                  
    
if __name__ == '__main__':
    main()
