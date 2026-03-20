"""
Serverless функция для генерации многолистового Excel отчета
URL: /api/download
УНИВЕРСАЛЬНАЯ ВЕРСИЯ - поддерживает английский, русский, казахский
"""

from http.server import BaseHTTPRequestHandler
import json
import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import io

# Конфигурация
KOBO_API_TOKEN = '929c90ea6bbce9e24789c10b2eb9740e3352d859'
ASSET_ID = 'aCE5fencfcUpVhvCRdCoxc'

# УНИВЕРСАЛЬНЫЙ маппинг городов
CITY_MAPPING = {
    # Английский
    'Astana': 'г. Астана',
    'Almaty': 'г. Алматы',
    'Shymkent': 'г. Шымкент',
    'Aktobe': 'Актобе Г.А.',
    # Русский
    'Астана': 'г. Астана',
    'Алматы': 'г. Алматы',
    'Шымкент': 'г. Шымкент',
    'Актобе': 'Актобе Г.А.',
    # Казахский
    'Астана қаласы': 'г. Астана',
    'Алматы қаласы': 'г. Алматы',
    'Шымкент қаласы': 'г. Шымкент',
    'Ақтөбе қаласы': 'Актобе Г.А.',
    # Старые коды
    '1': 'г. Астана',
    '2': 'г. Алматы',
    '3': 'г. Шымкент',
    '4': 'Актобе Г.А.'
}

# Готовность и согласие
WILLINGNESS_YES = [
    'Yes, willing to answer now',
    'Да, готов отвечать сейчас',
    'Иә, қазір жауап беруге дайынмын'
]

CONSENT_YES = [
    'Yes, I agree to participate',
    'Да, согласен(на) принять участие',
    'Да, соглашусь принять участие',
    'Иә, қатысуға келісемін'
]

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
        # Правильные названия полей из Kobo
        city_raw = record.get('City:', '')
        city = CITY_MAPPING.get(city_raw, city_raw)
        
        # Если город не распознан - частичное совпадение
        if city == city_raw and city_raw:
            for key, value in CITY_MAPPING.items():
                if key.lower() in city_raw.lower() or city_raw.lower() in key.lower():
                    city = value
                    break
        
        # Дата и время
        date_raw = record.get('Visit date:', '')
        time_raw = record.get('Visit time:', '')
        time_clean = str(time_raw).split('+')[0].split('.')[0] if time_raw else ''
        
        # Результат
        result = record.get('Visit result:', '')
        
        # Готовность и согласие
        willingness = record.get('Is the respondent willing to answer?', '')
        consent = record.get('Are you willing to participate in this survey?', '')
        q08 = record.get('q08_survey2', '')
        
        # Определяем завершенность
        is_completed = (
            willingness in WILLINGNESS_YES and
            consent in CONSENT_YES
        )
        
        # Категория
        if is_completed and str(q08).strip():
            q08_val = str(q08).strip()
            if q08_val == '1':
                category = 'Наемный работник'
            elif q08_val in ['2', '3', '4', '5']:
                category = 'Самозанятый/ИП'
            else:
                category = 'Другое'
        else:
            category = 'Другое'
        
        # Язык респондента
        language = record.get('Respondent language:', '')
        
        # Проверка на контакт
        is_contact = (
            'Contact established' in result if result else False or
            'Контакт установлен' in result if result else False
        )
        
        processed.append({
            'date': date_raw,
            'time': time_clean,
            'city': city,
            'peo': record.get('PEO number (electoral precinct)', ''),
            'segment': record.get('Segment number (1 to 5)', ''),
            'interviewer': record.get('Interviewer full name:', ''),
            'result': result,
            'category': category,
            'language': language,
            'attempt': record.get('Attempt number:', ''),
            'is_completed': is_completed,
            'is_contact': is_contact
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
    ws['E14'] = 'Наемных'
    ws['F14'] = 'Самозанятых'
    
    for col in ['A14', 'B14', 'C14', 'D14', 'E14', 'F14']:
        ws[col].font = header_font
        ws[col].fill = header_fill
    
    row = 15
    for city, quota in QUOTAS.items():
        city_completed = sum(1 for r in processed_data if r['city'] == city and r['is_completed'])
        city_employed = sum(1 for r in processed_data if r['city'] == city and r['category'] == 'Наемный работник')
        city_self = sum(1 for r in processed_data if r['city'] == city and r['category'] == 'Самозанятый/ИП')
        progress = round((city_completed / quota['total'] * 100) if quota['total'] > 0 else 0, 1)
        
        ws[f'A{row}'] = city
        ws[f'B{row}'] = city_completed
        ws[f'C{row}'] = quota['total']
        ws[f'D{row}'] = f"{progress}%"
        ws[f'E{row}'] = f"{city_employed}/{quota['employed']}"
        ws[f'F{row}'] = f"{city_self}/{quota['self_employed']}"
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
    """Создание листа по ПЕО"""
    ws = wb.create_sheet("Polling Station")
    
    headers = ['Город', 'ПЕО', 'Интервьюер', 'Всего визитов', 'Завершено', 'RR (%)', 'CR (%)']
    ws.append(headers)
    
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    
    # Группировка
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
    
    # Данные
    for (city, peo, interviewer), stats in sorted(peo_stats.items()):
        rr = round((stats['completed'] / stats['visits'] * 100) if stats['visits'] > 0 else 0, 1)
        cr = round((stats['contacts'] / stats['visits'] * 100) if stats['visits'] > 0 else 0, 1)
        
        ws.append([city, peo, interviewer, stats['visits'], stats['completed'], f"{rr}%", f"{cr}%"])
    
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
    
    headers = ['Интервьюер', 'Город', 'Всего визитов', 'Завершено', 'Неконтакты', 'Отказы', 'RR (%)', 'CR (%)']
    ws.append(headers)
    
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    
    # Группировка
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
        if 'refused' in record['result'].lower() or 'Отказ' in record['result']:
            int_stats[key]['refusals'] += 1
    
    # Данные
    for (interviewer, city), stats in sorted(int_stats.items()):
        non_contacts = stats['visits'] - stats['contacts']
        rr = round((stats['completed'] / stats['visits'] * 100) if stats['visits'] > 0 else 0, 1)
        cr = round((stats['contacts'] / stats['visits'] * 100) if stats['visits'] > 0 else 0, 1)
        
        ws.append([interviewer, city, stats['visits'], stats['completed'], non_contacts, stats['refusals'], f"{rr}%", f"{cr}%"])
    
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
    """Создание листа Raw Data"""
    ws = wb.create_sheet("Raw Data")
    
    headers = ['Дата', 'Время', 'Город', 'ПЕО', 'Сегмент', 'Интервьюер', 'Результат', 'Категория', 'Язык', 'Попытка']
    ws.append(headers)
    
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    
    # Данные
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
    """Создание Excel отчета"""
    wb = Workbook()
    
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    create_dashboard_sheet(wb, processed_data)
    create_peo_sheet(wb, processed_data)
    create_interviewer_sheet(wb, processed_data)
    create_raw_data_sheet(wb, processed_data)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output.getvalue()

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            records = fetch_kobo_data()
            processed_data = process_data(records)
            excel_bytes = create_excel_report(processed_data)
            
            filename = f"ENPF_Survey_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
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
