import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import os

credentials = ServiceAccountCredentials.from_json_keyfile_name('configs\\token.json')
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
client = gspread.authorize(credentials)


spreadsheet = client.open('Название таблицы')
def generate_Kolichestvo_strok():
    worksheet = spreadsheet.get_worksheet(0)
    data = worksheet.get_all_records()
    current_time = datetime.now()

    start_time = current_time - timedelta(days=1)
    start_time = start_time.replace(hour=7, minute=0, second=0, microsecond=0)
    end_time = current_time - timedelta(days=0)
    end_time = end_time.replace(hour=7, minute=0, second=0, microsecond=0)

    filtered_data = [record for record in data if start_time <= datetime.strptime(record['Дата'] + ' ' + record['Время'], '%d.%m.%Y %H:%M:%S') <= end_time]

    count_dict = {}
    for record in filtered_data:
        lastname = record['Ф.И.О.']
        if lastname in count_dict:
            count_dict[lastname] += 1
        else:
            count_dict[lastname] = 1


    df = pd.DataFrame(list(count_dict.items()), columns=['Ф.И.О.', 'Количество'])
    df = df.sort_values(by='Количество', ascending=False)

    plt.figure(figsize=(10, 6))
    bars = plt.barh(df['Ф.И.О.'], df['Количество'])

    for value, bar in zip(df['Количество'], bars):
        plt.text(value, bar.get_y() + bar.get_height() / 2, str(value), ha='left', va='center',fontname='Times New Roman')

    plt.xlabel('Количество',fontsize=12,fontname='Times New Roman')
    plt.title('Количество заполенных строк каждым дежурным', fontsize=16,fontname='Times New Roman')
    plt.gca().invert_yaxis()
    output_file1 = os.path.join('./output', 'Figure_2.png')
    plt.savefig(output_file1)

