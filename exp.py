import telebot
from telebot import types
import time

TOKEN = '6392689978:AAE2t1kOtLK6ZmeKRtfN-kcuBOF3SEFYRFs'
bot = telebot.TeleBot(TOKEN)


expenditures =['airfare(train)','lodging','per_diemKZ']

user_data = {}


@bot.message_handler(commands=['start'])
def start(message):
    user_data[message.from_user.id]={'airfare(train)':None,'lodging':None,'per_diemKZ':None}
    show_dropdown(message.chat.id)


def show_dropdown(chat_id):
    markup = types.ReplyKeyboardMarkup(row_width=1, resize_keyboard=True)
    for exp in expenditures:
        markup.add(types.KeyboardButton(exp))

    bot.send_message(chat_id, "Choose an option:", reply_markup=markup)


@bot.message_handler(func=lambda message: message.text in expenditures)
def handle_exp_selection(message):
    chat_id = message.chat.id
    selected_exp = message.text

    # Store the selected option in user_data or process it as needed
    # Ask the user to provide information for the selected option
    bot.send_message(chat_id, f"Provide information for {selected_exp}:")
    bot.register_next_step_handler(message, exp_info, selected_exp)


def exp_info(message,selected_exp):
    chat_id = message.chat.id
    
    user_data[message.chat.id][selected_exp]=message.text
    bot.send_message(chat_id, f"Your message:{user_data[message.chat.id][selected_exp]}:")
    #time.sleep(1)

    # After processing the information, show the dropdown again with unchosen options
    expenditures.remove(selected_exp)
    #if user_data[message.chat.id][selected_exp]:
    show_dropdown(message.chat.id)


bot.polling()