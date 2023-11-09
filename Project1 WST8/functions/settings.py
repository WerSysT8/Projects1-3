from dotenv import load_dotenv
import os
load_dotenv('../pythonProject/configs/.env')

SPREADSHEET_ID = os.getenv("MY_ID")
CHAT_ID = os.getenv("CHAT_ID")
BOT_TOKEN = os.getenv("BOT_TOKEN")
