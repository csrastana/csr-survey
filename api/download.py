"""
Serverless функция для генерации многолистового Excel отчета
URL: /api/download
"""

from http.server import BaseHTTPRequestHandler
import json
import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import io

# Конфигурация
KOBO_API_TOKEN = '929c90ea6bbce9e24789c10b2eb9740e3352d859'
ASSET_ID = 'aCE5fencfcUpVhvCRdCoxc'

CITY_MAPPING = {
    '1': 'г. Астана',
    '2': 'г. Алматы', 
    '3': 'г. Шымкент',
    '4': 'Актобе Г.А.'
}

QUOTAS = {
    'г. Астана': {'total': 800, 'employed': 608, 'self_employed': 192, 'peo_count': 32},
    'г. Алматы': {'total': 975, 'employed': 741, 'self_employed': 234, 'peo_count': 39},
    'г. Шымкент': {'total': 700, 'employed': 532, 'self_employed': 168, 'peo_count': 28},
    'Актобе Г.А.': {'total': 525, 'employed': 399, 'self_employed': 126, 'peo_count': 21}
}

def fetch_kobo_data():
    """Загрузка данных из Kobo"""
    url = f"https://kf.kobotoolbox.org/api/v2/assets/{ASSET_ID}/data.json"
    headers = {'Authorization': f'Token {KOBO_API_TOKEN}'}
    
    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    
    data = response.json()
    return data.get('results', [])

def process_data(records):
    """Обработка данных"""
    processed = []
    
    for record in records:
        city_code = record.get('city', '')
        city = CITY_MAPPING.get(str(city_code), 'Неизвестно')
        
        # Время и дата
        time_raw = record.get('group_xn8xb93/time', '')
        time_clean = str(time_raw).split('+')[0].split('.')[0] if time_raw else ''
        
        date_raw = record.get('group_xn8xb93/date', '')
        
        # Результат
        result = record.get('group_ip3jm92/result', '')
        
        # Категория
        willingness = record.get('willingness', '')
        consent = record.get('consent', '')
        q08 = record.get('q08_survey2', '')
        
        is_completed = (
            willingness == 'Да, готов отвечать сейчас' and
            consent in ['Да, согласен(на) принять участие']
        )
        
        if is_completed and str(q08).strip():
            if str(q08).strip() == '1':
                category = 'Наемный работник'
            elif str(q08).strip() in ['2', '3', '4', '5']:
                category = 'Самозанятый/ИП'
            else:
                category = 'Другое'
        else:
            category = 'Другое'
        
        processed.append({
            'date': date_raw,
            'time': time_clean,
            'city': city,
            'peo': record.get('group_xn8xb93/PEO', ''),
            'segment': record.get('group_xn8xb93/segment_num', ''),
            'interviewer': record.get('group_xn8xb93/int_name', ''),
            'result': result,
            'category': category,
            'language': record.get('group_xl1fx65/lang_resp', ''),
            'attempt': record.get('group_xn8xb93/attempt', ''),
            'is_completed': is_completed,
            'is_contact': result == 'Контакт установлен - дверь открыли'
        })
    
    return processed

def create_dashboard_sheet(wb, processed_data):
    """Создание листа Dashboard"""
    ws = wb.create_sheet("Dashboard", 0)
    
    # Стили
    header_font = Font(bold=True, size=14, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    
    # Заголовок
    ws['A1'] = 'Мониторинг полевого опроса ЕНПФ'
    ws['A1'].font = Font(bold=True, size=16)
    
    # Основные метрики
    total_visits = len(processed_data)
    completed = sum(1 for r in processed_data if r['is_completed'])
    contacts = sum(1 for r in processed_data if r['is_contact'])
    
    total_quota = sum(q['total'] for q in QUOTAS.values())
    rr = round((completed / total_visits * 100) if total_visits > 0 else 0, 1)
    cr = round((contacts / total_visits * 100) if total_visits > 0 else 0, 1)
    
    ws['A4'] = 'Метрика'
    ws['B4'] = 'Значение'
    ws['A4'].font = header_font
    ws['A4'].fill = header_fill
    ws['B4'].font = header_font
    ws['B4'].fill = header_fill
    
    metrics = [
        ('Всего визитов', total_visits),
        ('Завершено опросов', completed),
        ('Общая квота', total_quota),
        ('Прогресс по квоте (%)', f"{round((completed/total_quota*100) if total_quota > 0 else 0, 1)}%"),
        ('Response Rate (%)', f"{rr}%"),
        ('Contact Rate (%)', f"{cr}%")
    ]
    
    for idx, (metric, value) in enumerate(metrics, start=5):
        ws[f'A{idx}'] = metric
        ws[f'B{idx}'] = value
    
    # По городам
    ws['A13'] = 'Прогресс по городам'
    ws['A13'].font = Font(bold=True, size=12)
    
    ws['A14'] = 'Город'
    ws['B14'] = 'Завершено'
    ws['C14'] = 'Квота'
    ws['D14'] = 'Прогресс (%)'
    
    for col in ['A14', 'B14', 'C14', 'D14']:
        ws[col].font = header_font
        ws[col].fill = header_fill
    
    row = 15
    for city, quota in QUOTAS.items():
        city_completed = sum(1 for r in processed_data if r['city'] == city and r['is_completed'])
        progress = round((city_completed / quota['total'] * 100) if quota['total'] > 0 else 0, 1)
        
        ws[f'A{row}'] = city
        ws[f'B{row}'] = city_completed
        ws[f'C{row}'] = quota['total']
        ws[f'D{row}'] = f"{progress}%"
        row += 1
    
    # Автоширина
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

def create_peo_sheet(wb, processed_data):
    """Создание листа по ПЕО (Polling Station)"""
    ws = wb.create_sheet("Polling Station")
    
    # Заголовки
    headers = ['Город', 'ПЕО', 'Интервьюер', 'Всего визитов', 'Завершено', 'RR (%)', 'CR (%)']
    ws.append(headers)
    
    # Стили заголовков
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    
    # Группировка по ПЕО
    peo_stats = {}
    for record in processed_data:
        key = (record['city'], record['peo'], record['interviewer'])
        if key not in peo_stats:
            peo_stats[key] = {'visits': 0, 'completed': 0, 'contacts': 0}
        
        peo_stats[key]['visits'] += 1
        if record['is_completed']:
            peo_stats[key]['completed'] += 1
        if record['is_contact']:
            peo_stats[key]['contacts'] += 1
    
    # Добавление данных
    for (city, peo, interviewer), stats in sorted(peo_stats.items()):
        rr = round((stats['completed'] / stats['visits'] * 100) if stats['visits'] > 0 else 0, 1)
        cr = round((stats['contacts'] / stats['visits'] * 100) if stats['visits'] > 0 else 0, 1)
        
        ws.append([
            city,
            peo,
            interviewer,
            stats['visits'],
            stats['completed'],
            f"{rr}%",
            f"{cr}%"
        ])
    
    # Автоширина
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

def create_interviewer_sheet(wb, processed_data):
    """Создание листа по интервьюерам"""
    ws = wb.create_sheet("Enumerator & Supervisor")
    
    # Заголовки
    headers = ['Интервьюер', 'Город', 'Всего визитов', 'Завершено', 'Неконтакты', 'Отказы', 'RR (%)', 'CR (%)']
    ws.append(headers)
    
    # Стили
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    
    # Группировка по интервьюерам
    int_stats = {}
    for record in processed_data:
        key = (record['interviewer'], record['city'])
        if key not in int_stats:
            int_stats[key] = {'visits': 0, 'completed': 0, 'contacts': 0, 'refusals': 0}
        
        int_stats[key]['visits'] += 1
        if record['is_completed']:
            int_stats[key]['completed'] += 1
        if record['is_contact']:
            int_stats[key]['contacts'] += 1
        if record['result'] == 'Отказ домохозяйства':
            int_stats[key]['refusals'] += 1
    
    # Добавление данных
    for (interviewer, city), stats in sorted(int_stats.items()):
        non_contacts = stats['visits'] - stats['contacts']
        rr = round((stats['completed'] / stats['visits'] * 100) if stats['visits'] > 0 else 0, 1)
        cr = round((stats['contacts'] / stats['visits'] * 100) if stats['visits'] > 0 else 0, 1)
        
        ws.append([
            interviewer,
            city,
            stats['visits'],
            stats['completed'],
            non_contacts,
            stats['refusals'],
            f"{rr}%",
            f"{cr}%"
        ])
    
    # Автоширина
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

def create_raw_data_sheet(wb, processed_data):
    """Создание листа с сырыми данными"""
    ws = wb.create_sheet("Raw Data")
    
    # Заголовки
    headers = ['Дата', 'Время', 'Город', 'ПЕО', 'Сегмент', 'Интервьюер', 'Результат', 'Категория', 'Язык', 'Попытка']
    ws.append(headers)
    
    # Стили
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    
    # Добавление данных
    for record in processed_data:
        ws.append([
            record['date'],
            record['time'],
            record['city'],
            record['peo'],
            record['segment'],
            record['interviewer'],
            record['result'],
            record['category'],
            record['language'],
            record['attempt']
        ])
    
    # Автоширина
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width

def create_excel_report(processed_data):
    """Создание полного Excel отчета"""
    wb = Workbook()
    
    # Удаляем дефолтный лист
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # Создаем листы
    create_dashboard_sheet(wb, processed_data)
    create_peo_sheet(wb, processed_data)
    create_interviewer_sheet(wb, processed_data)
    create_raw_data_sheet(wb, processed_data)
    
    # Сохранение в память
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output.getvalue()

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            # Загрузка данных
            records = fetch_kobo_data()
            
            # Обработка
            processed_data = process_data(records)
            
            # Генерация Excel
            excel_bytes = create_excel_report(processed_data)
            
            # Имя файла
            filename = f"ENPF_Survey_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            # Ответ
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            
            self.wfile.write(excel_bytes)
            
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            
            error_response = {
                'error': str(e),
                'message': 'Ошибка генерации отчета'
            }
            self.wfile.write(json.dumps(error_response, ensure_ascii=False).encode('utf-8'))
