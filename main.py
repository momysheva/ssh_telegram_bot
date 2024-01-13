import os
import telebot
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import requests
from io import BytesIO
from googleapiclient.http import MediaIoBaseUpload
import gspread
from oauth2client.service_account import ServiceAccountCredentials

TELEGRAM_TOKEN = '6392689978:AAE2t1kOtLK6ZmeKRtfN-kcuBOF3SEFYRFs'
GOOGLE_DRIVE_CREDENTIALS_FILE = 'sshbot-401810-507d88f6b018.json'
EXISTING_FOLDER_ID = '1kPjjtGIHftqmDCMqe9WfU9R805-jir5M'
SPREADSHEET_NAME='Applications'

SPREADSHEET_WORKSHEET = 'A:A'

credentials = service_account.Credentials.from_service_account_file(
    GOOGLE_DRIVE_CREDENTIALS_FILE,
    scopes=['https://www.googleapis.com/auth/drive']
)
drive_service = build('drive', 'v3', credentials=credentials)

bot = telebot.TeleBot(TELEGRAM_TOKEN)

scope = ['https://www.googleapis.com/auth/spreadsheets']
sheets_credentials = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_DRIVE_CREDENTIALS_FILE, scope)
sheets_client = gspread.authorize(sheets_credentials)
spreadsheet = sheets_client.open(SPREADSHEET_NAME)
sheet = spreadsheet.worksheet(SPREADSHEET_WORKSHEET)

# Get the folder to store the spreadsheet
folder_query = f"'{EXISTING_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.spreadsheet'"
folder_results = drive_service.files().list(q=folder_query, fields='files(id, name)').execute()
spreadsheet_id = folder_results['files'][0]['id']  # Get the ID of the spreadsheet

# Open the spreadsheet using gspread
spreadsheet = sheets_client.open_by_key(spreadsheet_id)
sheet = spreadsheet.worksheet(SPREADSHEET_WORKSHEET)

@bot.message_handler(commands=['start'])
def start(message):
    bot.reply_to(message, 'Send me a file, and I will save it to your Google Drive folder!')

@bot.message_handler(content_types=['document'])
def handle_file(message):
        user_id = message.from_user.id
        user_name = message.from_user.username

        user_folder_name = f'{user_name}_folder'
        query = f"name='{user_folder_name}' and '{EXISTING_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder'"
        results = drive_service.files().list(q=query, fields='files(id)').execute()
        user_folders = results.get('files', [])

        if not user_folders:
            folder_metadata = {
                'name': user_folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [EXISTING_FOLDER_ID]
            }
            folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
            user_folder_id = folder['id']
        else:
            user_folder_id = user_folders[0]['id']

    
        file_id = message.document.file_id
        file_name = message.document.file_name
        file_info = bot.get_file(file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        file_metadata = {
        'name': message.document.file_name,
        'parents': [user_folder_id],  }

        media_body = MediaIoBaseUpload(BytesIO(downloaded_file), mimetype=message.document.mime_type, resumable=True)

        uploaded_file=drive_service.files().create(
        body=file_metadata,
        media_body=media_body
        ).execute()

        file_link = uploaded_file['webContentLink']  # Get the link of the uploaded file
        bot.reply_to(message, f'File "{file_name}" saved to your Google Drive folder!')
        
        sheet.append_row([user_name, file_name, file_link])


if __name__ == '__main__':
    bot.polling()

