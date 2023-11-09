from datetime import datetime, timedelta
from reportlab.lib.pagesizes import legal
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

def generate_pdf():
    pdfmetrics.registerFont(TTFont('Times New Roman', 'times.ttf'))


    output_folder = 'output'

    pdf_file = os.path.join(output_folder, "Thx.pdf")
    c = canvas.Canvas(pdf_file, pagesize=legal)
    page_width, page_height = legal

    c.setFont("Times New Roman", 12)

    main_title = 'Отчет заполнения базы данных "Мониторинг рисковых тем" в период'
    title_x = 80
    title_y = page_height - 25
    c.setFont("Times New Roman", 16)
    c.drawString(title_x, title_y, main_title)

    current_time = datetime.now()
    start_time = current_time - timedelta(days=1)
    start_time = start_time.replace(hour=7, minute=0, second=0, microsecond=0)
    end_time = current_time - timedelta(days=0)
    end_time = end_time.replace(hour=7, minute=0, second=0, microsecond=0)

    sub_title = f'с {start_time.strftime("%d.%m.%Y")} 7:00 по {end_time.strftime("%d.%m.%Y")} 7:00.'
    sub_title_x = 200
    sub_title_y = title_y - 20
    c.setFont("Times New Roman", 16)
    c.drawString(sub_title_x, sub_title_y, sub_title)

    photos = ['output\\Figure_1.png',
              'output\\Figure_2.png',
              'output\\results_table2.png']
    y = page_height - 450

    for photo in photos:
        img_width = page_width - 45
        img_height = img_width * 250 / 400
        x = 25
        c.drawImage(photo, x, y, width=img_width, height=img_height)
        y -= img_height - 10
        if os.path.exists(photo):
            os.remove(photo)

    c.showPage()
    c.save()

