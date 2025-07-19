import telebot

bot = telebot.TeleBot('7128425992:AAGbgXkXqUEzMTicL8Nv0Hgk8T2mst9G-sQ')

@bot.message_handler(commands=['start'])
def main(message):
    bot.send_message(message.chat.id, 'Заявка')

bot.polling(none_stop=True)