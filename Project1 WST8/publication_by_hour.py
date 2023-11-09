import gspread
from oauth2client.service_account import ServiceAccountCredentials
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import os


credentials = ServiceAccountCredentials.from_json_keyfile_name('configs\\token.json')
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
client = gspread.authorize(credentials)
spreadsheet = client.open('Название таблицы')


def generate_Kolichestvo_publicasi():
    worksheet = spreadsheet.get_worksheet(0)
    data = worksheet.get_all_records()
    current_time = datetime.now()

    start_time = current_time - timedelta(days=1)
    start_time = start_time.replace(hour=7, minute=0, second=0, microsecond=0)
    end_time = current_time - timedelta(days=0)
    end_time = end_time.replace(hour=7, minute=0, second=0, microsecond=0)

    hourly_counts = {str(i): 0 for i in range(24)}

    for record in data:
        date_str = record['Дата']
        time_str = record['Время']
        datetime_str = datetime.strptime(f"{date_str} {time_str}", "%d.%m.%Y %H:%M:%S")

        if start_time <= datetime_str < end_time:
            hour = (datetime_str.hour - 8) % 24
            hourly_counts[str(hour)] += 1

    hours = [str((hour + 8) % 24) for hour in range(24)]
    counts = list(hourly_counts.values())

    plt.figure(figsize=(10, 6))
    plt.plot(hours, counts, marker='o')
    plt.title('Количество публикаций по часам',fontsize=16, fontname='Times New Roman')
    plt.xlabel('Час',fontsize=12, fontname='Times New Roman')
    plt.ylabel('Количество публикаций',fontsize=12,fontname='Times New Roman')
    plt.grid(True)
    output_file = os.path.join('./output', 'Figure_1.png')
    plt.savefig(output_file)
