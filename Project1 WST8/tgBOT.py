import telebot
from functions.settings import CHAT_ID, BOT_TOKEN

bot = telebot.TeleBot(BOT_TOKEN)


def send_file():
    with open('output\\thx.pdf', 'rb') as document:
        bot.send_document(CHAT_ID, document)
