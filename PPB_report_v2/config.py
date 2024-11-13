from dotenv import load_dotenv
import os


load_dotenv('configs/.env')

SHEET_ID = os.getenv('SHEET_ID') #айди таблицы
FLODER_NAME = os.getenv('FLODER_NAME')# название листа
DRIVE_ID  = os.getenv('DRIVE_ID')# айди папки на гугл драйв
