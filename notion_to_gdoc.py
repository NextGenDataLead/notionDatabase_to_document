import os
import io
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/drive.file']

def authenticate_google_drive():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def upload_docx_to_gdoc(docx_file_path, gdoc_name):
    creds = authenticate_google_drive()
    try:
        service = build('drive', 'v3', credentials=creds)

        file_metadata = {
            'name': gdoc_name,
            'mimeType': 'application/vnd.google-apps.document'
        }
        media = MediaFileUpload(docx_file_path,
                                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                                resumable=True)
        
        file = service.files().create(body=file_metadata, media_body=media, fields='id,webViewLink').execute()
        print(f"Google Doc created: {file.get('name')} (ID: {file.get('id')})")
        print(f"View link: {file.get('webViewLink')}")
        return file.get('id'), file.get('webViewLink')

    except HttpError as error:
        print(f"An error occurred: {error}")
        print(f"Error details: {error.resp.status}, {error.resp.reason}")
        if error.resp.status == 403:
            print("Permission denied. Please ensure your Google Drive API scope includes write access.")
        return None, None

if __name__ == '__main__':
    # This part is for testing the script independently
    # In the main task, this will be called from notion_to_word.py or separately.
    docx_file = "Output/NotionContent.docx" # Assuming this file exists from the previous step
    gdoc_output_name = "NotionContent_GoogleDoc"

    if os.path.exists(docx_file):
        print(f"Attempting to upload {docx_file} and convert to Google Doc...")
        upload_docx_to_gdoc(docx_file, gdoc_output_name)
    else:
        print(f"Error: {docx_file} not found. Please run notion_to_word.py first to generate it.")
