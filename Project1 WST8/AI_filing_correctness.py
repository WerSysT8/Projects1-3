import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import os
from functions.settings import SPREADSHEET_ID


SERVICE_ACCOUNT_FILE = 'configs\\token.json'

credentials = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE,
                scopes=["https://www.googleapis.com/auth/spreadsheets"])
service = build('sheets', 'v4', credentials=credentials)
sheet = service.spreadsheets()
range_name_links = 'J2:J'
request_links = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID,
                                                    range=range_name_links)
response_links = request_links.execute()
values = response_links.get('values', [])
last_row = len(values) + 1

RANGE_NAME_ZAGOLOVOK = f'H2:H{last_row}'
RANGE_NAME_TEMA = f'G2:G{last_row}'
RANGE_NAME_RAYON = f'E2:E{last_row}'
RANGE_NAME_DATE = f'B2:B{last_row}'
RANGE_NAME_TIME = f'C2:C{last_row}'

request_date = sheet.get(spreadsheetId=SPREADSHEET_ID, ranges=RANGE_NAME_DATE,
                         fields='sheets.data.rowData.values.formattedValue')
response_date = request_date.execute()

request_time = sheet.get(spreadsheetId=SPREADSHEET_ID, ranges=RANGE_NAME_TIME,
                         fields='sheets.data.rowData.values.formattedValue')
response_time = request_time.execute()

request_zagolovok = sheet.get(spreadsheetId=SPREADSHEET_ID, ranges=RANGE_NAME_ZAGOLOVOK,
                              fields='sheets.data.rowData.values.effectiveFormat.backgroundColor')
response_zagolovok = request_zagolovok.execute()

request_tema = sheet.get(spreadsheetId=SPREADSHEET_ID, ranges=RANGE_NAME_TEMA,
                         fields='sheets.data.rowData.values.effectiveFormat.backgroundColor')
response_tema = request_tema.execute()

request_rayon = sheet.get(spreadsheetId=SPREADSHEET_ID, ranges=RANGE_NAME_RAYON,
                          fields='sheets.data.rowData.values.effectiveFormat.backgroundColor')
response_rayon = request_rayon.execute()

def generate_Tabl_Korekt_zapoln_neiron():
    def get_color(response_temp):
        result = []
        for sheet in response_temp['sheets']:
            for row in sheet['data'][0]['rowData']:
                for cell in row['values']:
                    color = cell['effectiveFormat']['backgroundColor']
                    result.append(color)
        return result


    data_zagolovok = get_color(response_zagolovok)
    data_tema = get_color(response_tema)
    data_rayon = get_color(response_rayon)

    data_zagolovok = get_color(response_zagolovok)
    data_tema = get_color(response_tema)
    data_rayon = get_color(response_rayon)
    data_date = []
    data_time = []

    for row_date, row_time in zip(response_date['sheets'][0]['data'][0]['rowData'],
                                  response_time['sheets'][0]['data'][0]['rowData']):
        cell_date = row_date['values'][0]
        cell_time = row_time['values'][0]
        data_date.append(cell_date.get('formattedValue', ''))
        data_time.append(cell_time.get('formattedValue', ''))

    df_zagolovok = pd.DataFrame({'Цвет': data_zagolovok})
    df_tema = pd.DataFrame({'Цвет': data_tema})
    df_rayon = pd.DataFrame({'Цвет': data_rayon})
    df_data = pd.DataFrame({'Дата': data_date})
    df_time = pd.DataFrame({'Время': data_time})

    dfh = pd.DataFrame({
        'Дата': data_date,
        'Время': data_time,
    })
    dfh['datetime'] = pd.to_datetime(dfh['Дата'] + ' ' + dfh['Время'],
                                        format='%d.%m.%Y %H:%M:%S')

    current_datetime = datetime.now()
    start_datetime = current_datetime - timedelta(days=1)
    start_datetime = start_datetime.replace(hour=7, minute=0,
                                    second=0, microsecond=0)

    end_datetime = current_datetime - timedelta(days=0)
    end_datetime = end_datetime.replace(hour=7, minute=0,
                                    second=0, microsecond=0)

    filtered_data = dfh[(dfh['datetime'] >= start_datetime) & (dfh['datetime'] <= end_datetime)]


    def determine_correctness(color):
        if color.get('red', 0) == 0 and color.get('green', 0) == 1 and color.get('blue', 0) == 0:
            return 'Зеленый'
        elif color.get('red', 0) > 0 and color.get('green', 0) > 0 and color.get('blue', 0) == 0:
            return 'Желтый'


    def process_dataframe(df, column_name):
        df['Корректность'] = df['Цвет'].apply(determine_correctness)
        correct_count = (df['Корректность'] == 'Зеленый').sum()
        incorrect_count = (df['Корректность'] == 'Желтый').sum()
        total_count = correct_count + incorrect_count
        if total_count == 0:
            percent_correct = 0
        else:
            percent_correct = round((correct_count / total_count) * 100, 2)

        return column_name, correct_count, \
               incorrect_count, total_count, f"{percent_correct}%"


    df_zagolovok_filtered = df_zagolovok[df_zagolovok.index.isin(filtered_data.index)].copy()
    df_tema_filtered = df_tema[df_tema.index.isin(filtered_data.index)].copy()
    df_rayon_filtered = df_rayon[df_rayon.index.isin(filtered_data.index)].copy()

    dataframes = [df_zagolovok_filtered, df_tema_filtered, df_rayon_filtered]

    results = []
    columns = ["Заголовок", "Тема", "Район", "Дата", "Время"]
    for df, column_name in zip(dataframes, ["Заголовок", "Тема", "Район"]):
        results.append(process_dataframe(df, column_name))

    results_df = pd.DataFrame(results, columns=["Поле", "Верно", "Неверно", "Всего(нейронка)", "% Верно", ])


    def calculate_total_cells(df):
        return len(df)


    results_df['Всего'] = [calculate_total_cells(df_zagolovok_filtered),
                           calculate_total_cells(df_tema_filtered),
                           calculate_total_cells(df_rayon_filtered)]

    fig, ax = plt.subplots(figsize=(10, 5))
    ax.axis('off')

    table_title = "Корректность заполнения нейронкой"
    ax.text(0.5, 0.64, table_title, fontsize=12, fontname='Times New Roman', ha='center', va='center')

    table_data = [results_df.columns.tolist()] + results_df.values.tolist()
    ax.table(cellText=table_data, cellLoc='center', loc='center')


    file_path = os.path.join('./output', 'results_table2.png')
    plt.savefig(file_path, bbox_inches='tight', pad_inches=0.01, dpi=400)

