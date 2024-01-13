import telebot
from telebot import types
import pandas as pd
import Constants as keys
import Responses as R

bot = telebot.TeleBot(keys.API_KEY)
try:
    df = pd.read_excel('applications.xlsx')
except FileNotFoundError:
    df = pd.DataFrame(columns=['email', 'major','motivationLetter'])
keyboard = types.ReplyKeyboardMarkup(row_width=2, one_time_keyboard=True)
yes_button = types.KeyboardButton("Yes")
no_button = types.KeyboardButton("No")
keyboard.add(yes_button, no_button)

@bot.message_handler(commands=["start"])
def start(message):
    doc = open("Guidelines.pdf", 'rb')
    bot.send_document(message.chat.id, doc)
    bot.send_message(message.chat.id, "Would you like to apply?", parse_mode='html')

@bot.message_handler(func=lambda message: message.text.lower() in ["yes", "no"])
def ask_for_email(message):
    selected_option = message.text.lower()
    if selected_option == "yes":
        bot.send_message(message.chat.id, "Please provide your email", parse_mode='html')
    else:
        bot.send_message(message.chat.id, "Okay, thank you!", parse_mode='html')
    

@bot.message_handler(func=lambda message: '@' in message.text)
def handle_email(message):
    user_email = message.text
    global df
    df = df.append({'email': user_email}, ignore_index=True)
    bot.send_message(message.chat.id, "Your email is recorded", parse_mode='html')
    bot.send_message(message.chat.id, "Please provide your Major", parse_mode='html')
    df.to_excel('applications.xlsx', index=False)

@bot.message_handler(func=lambda message:True)
def handle_major(message):
    user_major = message.text
    global df
    df.at[df.index[-1], 'major'] = user_major 
    bot.send_message(message.chat.id, "Your major is recorded", parse_mode='html')
    bot.send_message(message.chat.id, "Please provide your Year of Study", parse_mode='html')
    df.to_excel('applications.xlsx', index=False)



@bot.message_handler(content_types=['document'])
def handle_document(message):
    if message.document.mime_type == 'application/pdf':
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        # Save the PDF file
        file_path = f"{message.chat.id}_{message.document.file_name}"
        with open(file_path, 'wb') as new_file:
            new_file.write(downloaded_file)

        # Update the DataFrame with the file path
        global df
        df.at[df.index[-1], 'motivationLetter'] = file_path

        bot.send_message(message.chat.id, "Your letter is recorded", parse_mode='html')

        # Save the updated DataFrame to the Excel file
        df.to_excel('applications.xlsx', index=False)

    else:
        bot.send_message(message.chat.id, "Please upload a valid PDF file")


bot.polling(non_stop=True)
