import datetime
import io
from config import SHEET_ID, FLODER_NAME, DRIVE_ID
import re
from configparser import ConfigParser
from pathlib import Path
from datetime import date, timedelta
import gspread
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from oauth2client.service_account import ServiceAccountCredentials


BASE_DIR = Path(__file__).parent
config_ini = ConfigParser()
data_file = BASE_DIR / "config.ini"
config_ini.read(data_file)

# Настройка сервисного аккаунта и API Google Drive и Sheets
CREDENTIALS_FILE = './configs/reloadppb-134c53ec70ae.json'
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    ['https://www.googleapis.com/auth/spreadsheets',
     'https://www.googleapis.com/auth/drive',
     'https://www.googleapis.com/auth/drive.readonly',
     'https://www.googleapis.com/auth/documents.readonly']
)
service = build('drive', 'v3', credentials=credentials)
gc = gspread.authorize(credentials)
table = gc.open_by_key(SHEET_ID)
wks = table.worksheet(FLODER_NAME)

def get_last_date_from_sheet(worksheet):
    values = worksheet.col_values(1)
    last_date_str = values[-1] if values else None
    if last_date_str:
        last_date = datetime.datetime.strptime(last_date_str, "%d.%m.%Y").date()
        return last_date

# Сравнивает дату файла с датой из таблицы
def compare_dates(file_date, sheet_date):
    return file_date > sheet_date if sheet_date else True

# Скачивает все файлы .docx, новее даты в таблице, но не позже вчерашнего дня.
def download_docx_files(folder_id, last_sheet_date):
    yesterday = date.today() - timedelta(days=1)
    year = str(yesterday.year)
    month = f"{yesterday.month:02d}"
    date_prefix = f"{month}.{year}"
    print(f"Запрос списка файлов из корневой папки с ID: {folder_id}")
    results = service.files().list(q=f"'{folder_id}' in parents", pageSize=1000, fields="files(id, name)").execute()
    items = results.get('files', [])

    # Поиск папки за текущий год
    year_folder_id = next((item['id'] for item in items if item['name'] == year), None)
    if not year_folder_id:
        print(f"Папка за год {year} не найдена.")
        return []
    print(f"Папка за год {year} найдена, ID: {year_folder_id}")

    # Запрос списка файлов из папки за год
    results = service.files().list(q=f"'{year_folder_id}' in parents", pageSize=1000, fields="files(id, name)").execute()
    items = results.get('files', [])

    # Поиск папки за текущий месяц в формате MM.YYYY
    month_folder_id = next((item['id'] for item in items if item['name'] == date_prefix), None)
    if not month_folder_id:
        print(f"Папка за месяц {date_prefix} не найдена.")
        return []
    print(f"Папка за месяц {date_prefix} найдена, ID: {month_folder_id}")

    # Запрос списка файлов из папки за месяц
    results = service.files().list(q=f"'{month_folder_id}' in parents and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'",
                                    fields="files(id, name, createdTime)").execute()
    items = results.get('files', [])

    downloaded_files = []
    for item in items:
        file_date_str = item['createdTime']
        file_date = datetime.datetime.fromisoformat(file_date_str[:-1]).date()

        print(f"Дата файла: {file_date}, Дата из таблицы: {last_sheet_date}, Вчера: {yesterday}")

        if compare_dates(file_date, last_sheet_date) and file_date < (yesterday + timedelta(days=1)):
            file_id = item['id']
            file_name = BASE_DIR / 'PPB' / item['name']

            if not file_name.exists():
                print(f"Скачивание файла: {item['name']} (ID: {file_id})")
                request = service.files().get_media(fileId=file_id)
                with io.FileIO(file_name, 'wb') as fh:
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while not done:
                        status, done = downloader.next_chunk()
                        print(f"Загрузка {int(status.progress() * 100)}% завершена.")
                downloaded_files.append(file_name)
            else:
                print(f"Файл {file_name} уже существует, пропуск.")

    return downloaded_files

def parse_report_text(text, doc):
    data = []
    n = 0
    months_ = {
        'января': '01', 'февраля': '02', 'марта': '03', 'апреля': '04',
        'мая': '05', 'июня': '06', 'июля': '07', 'августа': '08',
        'сентября': '09', 'октября': '10', 'ноября': '11', 'декабря': '12'
    }

    list_info = re.split(r'Петр Бирюков|Пресс-служба КГХ', text)

    # Получаем ссылки сразу после обработки текста
    # link1 = []
    # rels1 = doc.part.rels
    # for rel1 in rels1:
    #     if rels1[rel1].reltype == RT.HYPERLINK:
    #         link1.append(rels1[rel1]._target)
    #





    last_date = None  # Инициализация last_date

    if len(list_info) > 1:
        for news in list_info[1:]:
            news_data = []

            # Извлечение даты для каждого блока новости
            date_string = re.search(r'(\d{1,2})\s+([а-я]+)\s+(\d{4})', list_info[n])
            n += 1

            if date_string:
                day_, month_str, year_ = date_string.groups()
                month_ = months_.get(month_str, '')
                date = f'{day_}.{month_}.{year_}' if month_ else ''
                news_data.append(date)
                last_date = date
            else:
                news_data.append(last_date if last_date else '')

            info = re.search(r'^(.*?)(?=–)', news)
            if info:
                news_data.append(info.group(0).strip(': '))
            else:
                news_data.append('')

            message_count = re.search(r'\d+\s*(?=сообщ|\sпубл)', news)
            news_data.append(int(message_count.group(0).strip()) if message_count else 'wrong')

            news_data.append(re.search(r'топ\s+яндекс', news, flags=re.IGNORECASE) is not None)

            fed_smi = re.search(r'(?<=Федеральные СМИ).*?(?=Городские СМИ)', news)
            news_data.append(fed_smi.group(0).strip(': ') if fed_smi else '')

            city_smi = re.search(r'(?<=Городские СМИ).*?(?=Бегущая)', news) or re.search(r'(?<=Городские СМИ).*', news)
            news_data.append(city_smi.group(0).strip(': ') if city_smi else '')

            tv = ', '.join([i.strip() for i in re.findall(r'(?<=Бегущ.. строк. на телеканале).*?(?=:)', news)])
            news_data.append(tv)

            facade = ', '.join([i.strip() for i in re.findall(r'(?<=Бегущ.. строк. на фасаде).*?(?=:)', news)])
            news_data.append(facade)


            data.append(news_data)




        links = []
        rels = doc.part.rels
        for rel in rels:
            if rels[rel].reltype == RT.HYPERLINK:
                links.append(rels[rel]._target)
        links.reverse()
        print(links)#вместо порядка ссылок "1 2 3" выводит "1 3 2 "

        if len(links) != len(data):
            raise ValueError(f"Количество ссылок ({len(links)}) не соответствует количеству блоков ({len(data)}).")


        for i in range(len(data)):
            data[i].append(links[i])
        links.clear()
        return data











def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list) + 1)

folder_id = DRIVE_ID
folder_path = BASE_DIR / 'PPB'
last_sheet_date = get_last_date_from_sheet(wks)

# Скачиваем все подходящие файлы
downloaded_files = download_docx_files(folder_id, last_sheet_date)


for downloaded_file in downloaded_files:
    print(f"Обработка файла: {downloaded_file}")
    try:
        doc = Document(downloaded_file)
        text = ' '.join(paragraph.text for paragraph in doc.paragraphs)

        # Очистка от спецсимволов
        remove_map = {
            ord('\n'): '',
            ord('\t'): '',
            ord('\b'): ''
        }
        text = text.translate(remove_map)

        parsed_data = parse_report_text(text, doc)
        print(parsed_data)




        for item in parsed_data:

            row = next_available_row(wks)
            cell_list = wks.range('A{}:I{}'.format(row, row))
            for cell, value in zip(cell_list, item):
                cell.value = value
            wks.update_cells(cell_list)



    except ValueError as ve:
        print(ve)
    except Exception as e:
        print(f"Ошибка при обработке файла {downloaded_file}: {e}")
    finally:
        downloaded_file.unlink()
        print(f"Файл {downloaded_file} был успешно удалён.")