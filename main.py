import telebot

bot = telebot.TeleBot ('7128425992:AAGbgXkXqUEzMTicL8Nv0Hgk8T2mst9G-sQ')

@bot.message_handlers(commands=['start'])
def main(message):
    bot.message_handlers(message.chat.id, 'Заявка')
