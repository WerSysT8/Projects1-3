import os
import logging
import asyncio
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from telegram import Bot
from dotenv import load_dotenv

load_dotenv()

SHEET_ID = os.getenv('SHEET_ID')
SHEET_TABS = os.getenv('SHEET_TABS', "Данные,Остальное").split(",")
TOKEN = os.getenv('TOKEN')
CHANNEL_ID = os.getenv('CHANNEL_ID')
GOOGLE_CREDENTIALS_FILE = os.getenv('GOOGLE_CREDENTIALS_FILE')

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO
)
logger = logging.getLogger(__name__)

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)

bot = Bot(token=TOKEN)

def get_sheet_data(sheet_name):
    try:
        sheet = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        data = sheet.get_all_records()
        print(f"Чтение данных из вкладки: {sheet_name} - Успешно")
        return data
    except Exception as e:
        logger.error(f"Ошибка при чтении вкладки {sheet_name}: {e}")
        return []

async def check_dates_and_notify():
    today = datetime.now().date()
    three_days_later = today + timedelta(days=3)

    urgent_notifications = []
    regular_notifications = []

    for sheet_name in SHEET_TABS:
        data = get_sheet_data(sheet_name)

        for row in data:
            if not row.get('Дата'):
                continue

            try:
                event_date = datetime.strptime(row['Дата'], '%d.%m.%Y').date()
            except ValueError as ve:
                logger.error(f"Ошибка обработки даты в строке: {row} - {ve}")
                continue

            if event_date == today:
                message = f"{row.get('Наименование')} ({row.get('ID')}) - {row.get('Страна')} - {event_date.strftime('%d.%m.%Y')}"
                regular_notifications.append(message)

            elif event_date == three_days_later:
                message = f"❗ {row.get('Наименование')} ({row.get('ID')}) - {row.get('Страна')} - {event_date.strftime('%d.%m.%Y')}"
                urgent_notifications.append(message)

    all_notifications = urgent_notifications + regular_notifications

    if all_notifications:
        final_message = "\n".join(all_notifications)
        try:
            await bot.send_message(chat_id=CHANNEL_ID, text=final_message)
            print("Все уведомления отправлены в одном сообщении.")
        except Exception as e:
            logger.error(f"Ошибка при отправке уведомлений: {e}")

def main():
    asyncio.run(check_dates_and_notify())
    print("Уведомления отправлены. Завершение работы.")

if __name__ == '__main__':
    main()

