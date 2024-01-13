import telebot
from telebot import types
import pandas as pd
import os 
import random
from google.oauth2 import service_account
from verification import send_verification_email
from pydrive.drive import GoogleDrive
from io import BytesIO
from googleapiclient.http import MediaIoBaseUpload
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import re
from apscheduler.schedulers.blocking import BlockingScheduler
import gspread

TOKEN = '6392689978:AAE2t1kOtLK6ZmeKRtfN-kcuBOF3SEFYRFs'
GOOGLE_DRIVE_CREDENTIALS_FILE = 'sshbot-401810-507d88f6b018.json'
EXISTING_FOLDER_ID = '1kPjjtGIHftqmDCMqe9WfU9R805-jir5M'
SPREADSHEET_ID = '1onQssJILUERje2KvacbM9rzSkCO5znoMqzpgzGRlLsY'
SHEET_NAME = 'Sheet1' 

credentials = service_account.Credentials.from_service_account_file(
    GOOGLE_DRIVE_CREDENTIALS_FILE,
    scopes=['https://www.googleapis.com/auth/drive']
)
drive_service = build('drive', 'v3', credentials=credentials)
bot = telebot.TeleBot(TOKEN)
drive:GoogleDrive

scope = ['https://www.googleapis.com/auth/spreadsheets']
credentials_sheets = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_DRIVE_CREDENTIALS_FILE, scope)
service = build('sheets', 'v4', credentials=credentials_sheets)

# Dictionary to store user data
user_data = {}
verification_codes={}
expenditures={}
expenditure={}
status={}


@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    doc = open("Guidelines.pdf", 'rb')
    bot.send_document(message.chat.id, doc)
    bot.reply_to(message, "Welcome! Type /apply to start the application process.")

@bot.message_handler(commands=['apply'])
def start_application(message):
    chat_id = message.chat.id
    user_id = message.from_user.id

    markup = types.InlineKeyboardMarkup()
    markup.row(types.InlineKeyboardButton("Yes", callback_data='yes'),
               types.InlineKeyboardButton("No", callback_data='no'))

    bot.send_message(chat_id, "Do you want to apply?", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: True)
def callback_query(call):
    chat_id = call.message.chat.id
    user_id = call.from_user.id
    if call.data == 'yes':
        bot.send_message(chat_id, "Great! Please provide the following information.\n\n1. Your email:")
        user_data[user_id] = {'drive_folder':None,'Number':None,'First Name': None,'Last Name': None,'ID Number': None,'Level of the Program': None,
                              'email': None,'Major': None, 'Year of study': None,'CGPA>=3':None,
                               'Covid Vaccination status':None,'step': 'motivation_letter', 'Trip destination':None,
                               'Trip Duration':None,'From':None,'To':None,'Purpose of the trip':None,'Faculty supervisor\'s name':None,
                               'Number of previously funded trip destinations (int\'l and domestic)':0,
                               'Amount of previously funded trips (in KZT)':0,'Funding source code':'','Link to the trip supporting documents':None, 'Currency':None
                               }
        expenditures[user_id]= {'airfare(train)':None,'lodging':None,'per_diem':None,'per_diemKZ':None,'medical_incurance':None,
                                                'local_transportation':None,'visa':None, 'registration_fee':None, 'membership':None,
                                                  'abstract_submission':None, 'Total':None, 'Total_KZT':None}
        status[user_id]={'row':None,'column':None}
        bot.register_next_step_handler_by_chat_id(chat_id, process_email, user_id)
        
    elif call.data == 'no':
        bot.send_message(chat_id, "Thank you! If you change your mind, type /apply to start the application process.")


def process_email(message, user_id):
    chat_id = message.chat.id
    email = message.text
    if re.match(r"[a-zA-Z0-9._%+-]+@nu\.edu\.kz", email):
        user_data[user_id]['email'] = email
        bot.send_message(message.chat.id, "Welcome! To verify your email, type /verify")
    else:
        bot.send_message(message.chat.id, "Invalid email address. Please provide an email from the @nu.edu.kz domain.")
        bot.register_next_step_handler_by_chat_id(chat_id, process_email, user_id)

def process_name(message, user_id):
    name=message.text
    user_data[user_id]['First Name'] = name
    bot.send_message(message.chat.id, "Please provide your Surname")
    bot.register_next_step_handler_by_chat_id(message.chat.id, process_surname, user_id)


def process_surname(message, user_id):
    surname=message.text
    user_data[user_id]['Last Name'] = surname
    bot.send_message(message.chat.id, "Please provide your Student ID")
    bot.register_next_step_handler_by_chat_id(message.chat.id, process_studentID, user_id)

def process_studentID(message, user_id):
    studentID=message.text
    user_data[user_id]['ID Number'] = studentID
    markup = types.ReplyKeyboardMarkup(row_width=1,resize_keyboard=True)
    option1 = types.KeyboardButton("Undergraduate")
    option2 = types.KeyboardButton("Graduate")
    option3 = types.KeyboardButton("PhD")
    markup.add(option1, option2, option3)

    bot.send_message(message.chat.id, "Please select your level of the program:", reply_markup=markup)
    bot.register_next_step_handler(message, process_program_level)


def process_program_level(message):
    user_id = message.from_user.id
    user_data[user_id]['Level of the Program'] = message.text
    if message.text == 'Undergraduate':
        ask_major(message, ['Political Science', 'Sociology', 'Anthropology', 'History', 'Economics', 'WLL', 'Biological Sciences', 'Chemistry', 'Mathematics', 'Physics'])
    elif message.text == 'Graduate':
        ask_major(message, ['Political Science', 'Economics', 'Eurasian Studies', 'Biological Sciences', 'Chemistry', 'Applied Mathematics', 'Physics'])
    elif message.text == 'PhD':
        ask_major(message, ['Life Sciences', 'Eurasian Studies', 'Chemistry', 'Mathematics', 'Physics'])


def ask_major(message, majors):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for major in majors:
        markup.add(types.KeyboardButton(major))

    bot.send_message(message.chat.id, "Please select your major:", reply_markup=markup)
    bot.register_next_step_handler(message, process_major)

def process_major(message):
    user_id = message.from_user.id
    user_data[user_id]['Major'] = message.text
    if user_data[user_id]['Level of the Program'] == 'Undergraduate':
        markup = types.ReplyKeyboardMarkup(row_width=1,resize_keyboard=True)
        option1 = types.KeyboardButton("1")
        option2 = types.KeyboardButton("2")
        option3 = types.KeyboardButton("3")
        option4 = types.KeyboardButton("4")
        markup.add(option1, option2, option3,option4)
        bot.send_message(message.chat.id, "Please provide your Year of Study", reply_markup=markup)
        bot.register_next_step_handler(message, process_year_of_study)

    elif user_data[user_id]['Level of the Program'] == 'Graduate':
        markup = types.ReplyKeyboardMarkup(row_width=1,resize_keyboard=True)
        option1 = types.KeyboardButton("1")
        option2 = types.KeyboardButton("2")
        markup.add(option1, option2)
        bot.send_message(message.chat.id, "Please provide your Year of Study", reply_markup=markup)
        bot.register_next_step_handler(message, process_year_of_study)
    elif user_data[user_id]['Level of the Program'] == 'PhD':
        markup = types.ReplyKeyboardMarkup(row_width=1,resize_keyboard=True)
        option1 = types.KeyboardButton("1")
        option2 = types.KeyboardButton("2")
        option3 = types.KeyboardButton("3")
        option4 = types.KeyboardButton("4")
        markup.add(option1, option2, option3,option4)
        bot.send_message(message.chat.id, "Please provide your Year of Study", reply_markup=markup)
        bot.register_next_step_handler(message, process_year_of_study)

    
def process_year_of_study(message):
    user_id = message.from_user.id
    reply_markup = types.ReplyKeyboardRemove()
    user_data[user_id]['Year of study'] = message.text
    bot.send_message(message.chat.id, "Please provide your CGPA", reply_markup=reply_markup)
    bot.register_next_step_handler(message, process_cgpa)

def process_cgpa(message):
    user_id = message.from_user.id
    user_data[user_id]['CGPA>=3'] = message.text
    markup = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    option1 = types.KeyboardButton("G1:Fully vaccinated")
    option2 = types.KeyboardButton("G2:Individuals with natural immunity")
    option3 = types.KeyboardButton("G3: Not fully vaccinated")
    option4 = types.KeyboardButton("G4: Not vaccinated due to medical exemptions")
    option5 = types.KeyboardButton("G5: Not vaccinated")
    markup.add(option1, option2, option3,option4,option5)
    bot.send_message(message.chat.id, "Please select your Covid Vaccination Status:", reply_markup=markup)
    bot.register_next_step_handler(message, process_covid_status)

def remove_word(input_string, word_to_remove):
    words = input_string.split()  # Split the string into a list of words
    filtered_words = [word for word in words if word.lower() != word_to_remove.lower()]
    result_string = ' '.join(filtered_words)  # Join the remaining words back into a string
    return result_string

def show_dropdown(message,expenditure):
    user_id = message.from_user.id
    chat_id=message.chat.id
    if expenditure:
        markup = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
        for exp in expenditure:
            markup.add(types.KeyboardButton(exp))
        bot.send_message(chat_id, "Choose an option:", reply_markup=markup)
        bot.register_next_step_handler(message, handle_exp_selection)
    else:
        expenditures[user_id]['Total'] = (
            int(remove_word(expenditures[user_id]['airfare(train)'], user_data[user_id]['Currency']))
            + int(remove_word(expenditures[user_id]['lodging'], user_data[user_id]['Currency']))
            + int(remove_word(expenditures[user_id]['per_diem'], user_data[user_id]['Currency']))
            + int(remove_word(expenditures[user_id]['medical_incurance'], user_data[user_id]['Currency']))
            + int(remove_word(expenditures[user_id]['local_transportation'], user_data[user_id]['Currency']))
            + int(remove_word(expenditures[user_id]['visa'], user_data[user_id]['Currency']))
            + int(remove_word(expenditures[user_id]['registration_fee'], user_data[user_id]['Currency']))
            + int(remove_word(expenditures[user_id]['membership'], user_data[user_id]['Currency']))
            + int(remove_word(expenditures[user_id]['abstract_submission'], user_data[user_id]['Currency']))
        )

        expenditures[user_id]['Total_KZT']=int(remove_word(expenditures[user_id]['per_diemKZ'], 'KZT'))
        save_to_excel(user_id)
        user_data[user_id].clear()
        expenditures[user_id].clear()

def process_covid_status(message):
    reply_markup = types.ReplyKeyboardRemove()
    print('covid')
    user_id = message.from_user.id
    user_data[user_id]['Covid Vaccination status'] = message.text
    bot.send_message(message.chat.id, "Your Covid Vaccination Status is saved.\n Please provide your Trip Destination", reply_markup=reply_markup)
    bot.register_next_step_handler(message, trip_destination)

    
def trip_destination(message):
    user_id = message.from_user.id
    user_data[user_id]['Trip destination'] = message.text
    bot.send_message(message.chat.id, "Please provide your Trip Duration (in days)")
    bot.register_next_step_handler(message, process_trip_duration)

def process_trip_duration(message):
    user_id = message.from_user.id
    user_data[user_id]['Trip Duration'] = message.text
    bot.send_message(message.chat.id, "Please provide starting date of your Trip Period(DD/MM/YYYY)")
    bot.register_next_step_handler(message, process_trip_period_from)

def process_trip_period_from(message):
    user_id = message.from_user.id
    # user_data[user_id]['Trip period'] = message.text
    user_data[user_id]['From']=message.text
    bot.send_message(message.chat.id, "Please provide ending date of your Trip Period(DD/MM/YYYY)")
    bot.register_next_step_handler(message, process_trip_period_to)

def process_trip_period_to(message):
    user_id = message.from_user.id
    user_data[user_id]['To']=message.text
    # user_data[user_id]['Trip period']=trip_period[user_id]
    bot.send_message(message.chat.id, "Please provide your Purpose of the trip")
    bot.register_next_step_handler(message, trip_purpose)

def trip_purpose(message):
    user_id = message.from_user.id
    user_data[user_id]['Purpose of the trip'] = message.text
    bot.send_message(message.chat.id, "Please provide your Faculty Supervisor's Name")
    bot.register_next_step_handler(message, process_supervisor)

def process_supervisor(message):
    user_id = message.from_user.id
    user_data[user_id]['Faculty supervisor\'s name'] = message.text
    bot.send_message(message.chat.id, "Your information is saved. Please provide your Motivation Letter (SEND AS A DOCUMENT)")
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)

def handle_user_input_after_supervisor(message):
    # Check if the user sent a document
    if message.content_type == 'document':
        handle_file(message)
    else:
        # Handle the case when the user types something instead of sending a document
        bot.send_message(message.chat.id, "Wrong format. Please upload your file (SEND AS A DOCUMENT)")
        bot.register_next_step_handler(message, handle_user_input_after_supervisor)

def handle_exp_selection(message):
    chat_id = message.chat.id
    selected_exp = message.text
    reply_markup = types.ReplyKeyboardRemove()
    bot.send_message(chat_id, f"Provide information for {selected_exp}:",reply_markup=reply_markup)
    bot.register_next_step_handler(message, exp_info, selected_exp)

def exp_info(message,selected_exp):
    chat_id = message.chat.id
    user_id = message.from_user.id
    if selected_exp!='per_diemKZ':
        expenditures[user_id][selected_exp]=message.text+' '+user_data[user_id]['Currency']
    else:
        expenditures[user_id][selected_exp]=message.text+' KZT'
    reply_markup = types.ReplyKeyboardRemove()
    #user_data[user_id]['expenditures'][selected_exp]=message.text
    #print(user_data[user_id]['expenditures'][selected_exp])
    #expenditure_limits[user_id][selected_exp]=message.text
    bot.send_message(chat_id, "Your information is saved!", reply_markup=reply_markup)
    #time.sleep(1)
    # After processing the information, show the dropdown again with unchosen options
    expenditure[user_id].remove(selected_exp)
    #if user_data[message.chat.id][selected_exp]:
    show_dropdown(message,expenditure[user_id])



@bot.message_handler(content_types=['document'])
def handle_file(message):
        user_id = message.from_user.id
        current_step = user_data[user_id].get('step')
        if current_step == 'motivation_letter':
            handle_motivation_letter(message,user_id)
        elif current_step == 'request_form':
            handle_student_travel_funding_request_form(message,user_id)
        elif current_step == 'trip_form':
            handle_trip_form(message,user_id)
        elif current_step == 'trip_form_expenses':
            handle_travel_form_expenses(message,user_id)
        elif current_step == 'recommendation_form':
            handle_recommendation_form(message,user_id)
        elif current_step == 'absense_form':
            handle_absense_form(message,user_id)
        elif current_step == 'letter_of_invitation':
            handle_letter_of_invitation(message,user_id)
        elif current_step == 'covid_passport':
            handle_covid_passport(message,user_id)
        elif current_step == 'passport_id':
            handle_passport_id(message,user_id)
        elif current_step == 'registration_fee':
            handle_registration_fee(message,user_id)
        elif current_step == 'lodging_invoice':
            handle_lodging_invoice(message,user_id)
        elif current_step == 'tickets_info':
            handle_tickets_info(message,user_id)
        else:
            bot.reply_to(message, "Unexpected file upload. Please follow the instructions.")
        
        
def handle_motivation_letter(message, user_id):
    file_to_drive_folder(message,user_id)
    user_data[user_id]['step']='request_form'
    bot.reply_to(message, 'Your motivation letter is saved!\n\nPlease upload STUDENT TRAVEL FUNDING REQUEST')
    doc = open("1. BLANK STUDENT TRAVEL FUNDING REQUEST.docx", 'rb')
    bot.send_document(message.chat.id, doc)
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)
        
    

def handle_student_travel_funding_request_form(message, user_id):
    file_to_drive_folder(message,user_id)
    user_data[user_id]['step']='trip_form'
    bot.reply_to(message, 'Your STUDENT TRAVEL FUNDING REQUEST is saved!\n\nPlease upload Trip Report')
    doc = open("Trip report_eng.docx", 'rb')
    bot.send_document(message.chat.id, doc)
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)

def handle_trip_form(message, user_id):
    file_to_drive_folder(message,user_id)
    user_data[user_id]['step']='trip_form_expenses'
    bot.reply_to(message, 'Your Trip Report is saved!\n\nPlease upload Trip Report on Actual Expenses')
    doc = open("Trip report_on actual expenses.docx", 'rb')
    bot.send_document(message.chat.id, doc)
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)

def handle_travel_form_expenses(message, user_id):
    file_to_drive_folder(message,user_id)
    user_data[user_id]['step']='recommendation_form'
    bot.reply_to(message, 'Your Trip Report on actual expenses is saved!\n\nPlease upload Completed and signed by faculty member Blank SSH Faculty recommendation form for student travel')
    doc = open("2.BLANK SSH Faculty Recommendation Form.docx", 'rb')
    bot.send_document(message.chat.id, doc)
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)

def handle_recommendation_form(message, user_id):
    file_to_drive_folder(message,user_id)
    bot.reply_to(message, 'Your Recommendation Form is saved!\n\nPlease upload Absence from classes form signed by the professors and the Dean')
    user_data[user_id]['step']='absense_form'
    doc = open("3. BLANK Absence form.docx", 'rb')
    bot.send_document(message.chat.id, doc)
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)

def handle_absense_form(message, user_id):
    file_to_drive_folder(message,user_id)
    bot.reply_to(message, 'Your Absense Form is saved!\n\nPlease upload Official Letter of invitation to event/conference with a name of the invited student and confirming the students role at the event (i.e., presenter, delegate, or debater for example')
    user_data[user_id]['step']='letter_of_invitation'
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)


def handle_letter_of_invitation(message,user_id):
    file_to_drive_folder(message,user_id)
    bot.reply_to(message, 'Your Letter of Invitation is saved!\n\nPlease upload your scanned version of the Covid Vaccination passport')
    user_data[user_id]['step']='covid_passport'
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)
    
def handle_covid_passport(message, user_id):
    file_to_drive_folder(message,user_id)
    user_data[user_id]['step']='passport_id'
    bot.reply_to(message, 'Your scanned version of the Covid Vaccination passport is saved!\n\nPlease upload A scanned version of the international passport/state ID')
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)

def handle_passport_id(message, user_id):
    file_to_drive_folder(message,user_id)
    user_data[user_id]['step']='registration_fee'
    bot.reply_to(message, 'Your scanned version of the international passport/state ID is saved!\n\nPlease upload Registration fee invoice')
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)

def handle_registration_fee(message, user_id):
    file_to_drive_folder(message,user_id)
    user_data[user_id]['step']='lodging_invoice'
    bot.reply_to(message, 'Your Registration fee invoice is saved!\n\nPlease upload Lodging invoice')
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)

def handle_lodging_invoice(message, user_id):
    file_to_drive_folder(message,user_id)
    user_data[user_id]['step']='tickets_info'
    bot.reply_to(message, 'Your Lodging invoice is saved!\n\nPlease upload Air/train tickets booking indicating a preferable trip route, dates and visa fee information provided by Travel Agency')
    bot.reply_to(message, '''
    traveltse@bstravel.kz, +7 701 302 2370 (Sandugash - registration fee payment proceedings and lodging booking)
    visatse@bstravel.kz, +7 701 795 1045 (Zhansulu - visa proceedings)
    24@bstravel.kz, university@bstravel.kz +7 701 981 4204 (Zarina - air/train tickets booking and buying)
    ''')
    bot.register_next_step_handler(message, handle_user_input_after_supervisor)

def handle_tickets_info(message, user_id):
    file_to_drive_folder(message,user_id)
    user_data[user_id]['step']='completed'
    bot.reply_to(message, 'Air/train tickets booking indicating a preferable trip route, dates and visa fee information is saved')
    bot.send_message(message.chat.id, "Please provide information about your expenditures (DOLLAR/EURO Depending on your trip destination).If there is no need for particular expenditure type please INDICATE it as 0. Write c")
    markup = types.ReplyKeyboardMarkup(row_width=2, resize_keyboard=True)
    option_usd = types.KeyboardButton("USD")
    option_eur = types.KeyboardButton("EUR")
    markup.add(option_usd, option_eur)

    # Send the currency selection message with the custom keyboard
    bot.send_message(message.chat.id, "Please select the currency (USD or EUR):", reply_markup=markup)
    
    
    
    # expenditure[user_id] = [
    # 'airfare(train)', 'lodging', 'per_diem', 'per_diemKZ', 'medical_incurance',
    # 'local_transportation', 'visa', 'registration_fee', 'membership', 'abstract_submission']
    # bot.register_next_step_handler(message, show_dropdown,expenditure[user_id])
    # # save_to_excel(user_id)
    # # user_data.clear()
    bot.register_next_step_handler(message, handle_currency_selection, user_id)


def handle_currency_selection(message, user_id):
    user_currency = message.text.upper()  # Convert the user's input to uppercase
    if user_currency in ['USD', 'EUR']:
        # Save the user's selected currency in the user_data dictionary
        user_data[user_id]['Currency'] = user_currency

        # Continue with the expenditures dropdown
        expenditure[user_id] = [
            'airfare(train)', 'lodging', 'per_diem', 'per_diemKZ', 'medical_incurance',
            'local_transportation', 'visa', 'registration_fee', 'membership', 'abstract_submission'
        ]
        show_dropdown(message, expenditure[user_id])

    else:
        # Handle the case when the user provides an invalid currency
        bot.send_message(message.chat.id, "Invalid currency. Please select either USD or EUR.")
        bot.register_next_step_handler(message, handle_currency_selection, user_id)


def file_to_drive_folder(message,user_id):
    file_id = message.document.file_id
    file_name = message.document.file_name
    file_info = bot.get_file(file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    file_metadata = {
    'name': message.document.file_name,
    'parents': [user_data[user_id]['drive_folder']],  }

    media_body = MediaIoBaseUpload(BytesIO(downloaded_file), mimetype=message.document.mime_type, resumable=True)

    uploaded_file=drive_service.files().create(
    body=file_metadata,
    media_body=media_body
    ).execute()


def save_to_excel(applicant):
    df = pd.DataFrame.from_dict(user_data[applicant], orient='index')
    expenditures_data = expenditures[applicant]
    if expenditures_data:
        first_key, first_value = next(iter(expenditures_data.items()))
    df = df.T
    columns_to_remove = ['drive_folder', 'email', 'step', 'Currency']
    df = df.drop(columns=columns_to_remove, errors='ignore')
    try:
        # Get the existing data from Google Sheets
        result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='A:Z').execute()
        values = result.get('values', [])

        # If the spreadsheet is empty, create a new sheet
        if not values:
            values = [df.columns.tolist()]  # Add headers if the sheet is empty
            values[0] += ['Expenditures Request', 'Requested amount for school funding (in KZT), adjusted based on the limits', 'Member present', 'Vote','Comments (if needed)']

        last_row_index = len(values)
        next_empty_row_for_merging = last_row_index 
        print(last_row_index)
        # Append new data as a new row to the existing data
        combined_row = df.values.tolist()[0] + [first_key, first_value]
        values.append(combined_row)
        #values.append(df.values.tolist()[0])  # Assuming there's only one row to insert

        # Write the updated data back to the Google Sheet
        request = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range='A1',  # Start from the first row
            valueInputOption='RAW',
            body={'values': values}
        )
        response = request.execute()
        expenditure_excel(applicant, next_empty_row_for_merging)
        bot.send_message(applicant, "Thank you! Your application has been submitted.")
    except Exception as e:
        print(f"Error: {e}")
   
def expenditure_excel(applicant, next_empty_row_for_merging):
    expenditures_data = expenditures[applicant]
    if expenditures_data:
    # Remove the first key-value pair
        first_key, first_value = next(iter(expenditures_data.items()))
        del expenditures_data[first_key]
    try:
        # Get the existing data from Google Sheets
        result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range='A:Z').execute()
        values = result.get('values', [])
        # Iterate through all items in expenditures_data and append them as new rows
        for key, value in expenditures_data.items():
            # Find the index of 'Expenditure' and 'Values' columns
            expenditure_index = values[0].index('Expenditures Request')
            values_index = values[0].index('Requested amount for school funding (in KZT), adjusted based on the limits')

            new_row = [None] * len(values[0])
            new_row[expenditure_index] = key
            new_row[values_index] = value

            values.append(new_row)

        # Write the updated data back to the Google Sheet
        request = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range='A1',  # Start from the first row
            valueInputOption='RAW',
            body={'values': values}
        )
        response = request.execute()
        spreadsheet_data = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        first_sheet_id = spreadsheet_data['sheets'][0]['properties']['sheetId']
        merge_request = merge_cells_vertically(first_sheet_id, next_empty_row_for_merging, 0,next_empty_row_for_merging+len(expenditures_data)+1, len(user_data[applicant])-4)
        service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={'requests': [merge_request]}).execute()
        status[applicant]['row']=len(expenditures_data)+1
        status[applicant]['column']=len(user_data[applicant])-4
        print('Merged')

    except Exception as e:
        print(f"Error: {e}")

def merge_cells_vertically(sheet_id, start_row, start_col, end_row, end_col):
    merge_request = {
        "mergeCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,  
                "endRowIndex": end_row,
                "startColumnIndex": start_col,  
                "endColumnIndex": end_col
            },
            "mergeType": "MERGE_COLUMNS"
        }
    }
    return merge_request

@bot.message_handler(commands=['checkstatus'])
def check_status(message):
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_DRIVE_CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)

    # Open the spreadsheet by title
    spreadsheet_title = 'Applications'
    spreadsheet = client.open(spreadsheet_title)

    # Specify the sheet and cells to check
    sheet_name = 'Лист1'
    start_row = status[message.from_user.id]['row']
    column = status[message.from_user.id]['column']
    cells_to_check = [start_row ]

    sheet = spreadsheet.worksheet(sheet_name)
    cell_value = sheet.acell(cells_to_check).value

    if cell_value:
        print(f"Data exists in {cells_to_check}: {cell_value}")
        bot.send_message(message.chat.id, f"Your application status: {cell_value}")
    else:
        print(f"No data found in {cells_to_check}")
        bot.send_message(message.chat.id, "No data found. Please try again later.")

        # Set up the scheduler
    scheduler = BlockingScheduler()

    # Schedule the job to run every minute
    scheduler.add_job(cells_to_check, 'interval', minutes=1)

    # Start the scheduler
    try:
        scheduler.start()
    except KeyboardInterrupt:
        pass

@bot.message_handler(commands=['verify'])
def verify_email(message):
    user_email = user_data[message.from_user.id]['email']  
    verification_code = str(random.randint(100000, 999999))
    
    # Store the verification code (temporary)
    verification_codes[user_email] = verification_code
    
    # Send the verification email
    send_verification_email(user_email, verification_code)
    
    bot.send_message(message.chat.id, "Verification code sent to your email. Enter the code to verify.")
    bot.register_next_step_handler(message, handle_verification_code)

#@bot.message_handler(func=lambda message: True)
def handle_verification_code(message):
    user_email = user_data[message.from_user.id]['email']  # Replace with the actual user's email
    user_input = message.text.strip()
    
    if user_input == verification_codes.get(user_email):
        # Code is correct, mark email as verified in your database
        user_folder_name = f'{user_email}_folder'
        query = f"name='{user_folder_name}' and '{EXISTING_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder'"
        results = drive_service.files().list(q=query, fields='files(id)').execute()
        user_folders = results.get('files', [])

        if not user_folders:
            folder_metadata = {
                'name': user_folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [EXISTING_FOLDER_ID]
            }
            folder = drive_service.files().create(body=folder_metadata, fields='id,webViewLink').execute()
            user_folder_id = folder['id']
        else:
            user_folder_id = user_folders[0]['id']
            folder=drive_service.files().get(fileId=user_folder_id, fields='webViewLink').execute()

        user_data[message.from_user.id]['Link to the trip supporting documents']=folder.get('webViewLink', '')
        user_data[message.from_user.id]['drive_folder']=user_folder_id
        bot.send_message(message.chat.id, "Email verified successfully!")
        bot.send_message(message.chat.id, "Please provide your Name")
        bot.register_next_step_handler_by_chat_id(message.chat.id, process_name, message.from_user.id)
    else:
        bot.send_message(message.chat.id, "Invalid verification code. Try again.")


if __name__ == "__main__":
    bot.polling(none_stop=True)
