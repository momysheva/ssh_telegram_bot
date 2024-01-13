import os
import telebot
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# Set up Google Drive authentication
def authenticate_google_drive():
    gauth = GoogleAuth()
    gauth.LocalWebserverAuth()

    if not gauth.credentials:
        gauth.LocalWebserverAuth()
    elif gauth.access_token_expired:
        gauth.Refresh()
    else:
        gauth.Authorize()

    gauth.SaveCredentialsFile("sshbot-401810-507d88f6b018.json")
    drive = GoogleDrive(gauth)
    return drive

# Initialize Telebot and authenticate Google Drive
TOKEN = "6392689978:AAE2t1kOtLK6ZmeKRtfN-kcuBOF3SEFYRFs"
bot = telebot.TeleBot(TOKEN)
drive = authenticate_google_drive()

# Handler for documents
@bot.message_handler(content_types=['document'])
def handle_documents(message):
    try:
        # Download the document
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        file_path = f"downloads/{message.document.file_name}"

        # Save to Google Drive
        drive_file = drive.CreateFile({'title': message.document.file_name})
        drive_file.Upload()

        # Cleanup: Delete the downloaded file
        os.remove(file_path)

    except Exception as e:
        bot.reply_to(message, f"An error occurred: {str(e)}")

# Start the bot
bot.polling(none_stop=True)
