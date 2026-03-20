"""
Serverless функция для получения статистики опроса
URL: /api/stats
УНИВЕРСАЛЬНАЯ ВЕРСИЯ - поддерживает английский, русский, казахский
"""

from http.server import BaseHTTPRequestHandler
import json
import requests
from datetime import datetime

# Конфигурация
KOBO_API_TOKEN = '929c90ea6bbce9e24789c10b2eb9740e3352d859'
ASSET_ID = 'aCE5fencfcUpVhvCRdCoxc'

# УНИВЕРСАЛЬНЫЙ маппинг городов (английский, русский, казахский)
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
    # Казахский (если используется)
    'Астана қаласы': 'г. Астана',
    'Алматы қаласы': 'г. Алматы',
    'Шымкент қаласы': 'г. Шымкент',
    'Ақтөбе қаласы': 'Актобе Г.А.',
    # Старые коды (на всякий случай)
    '1': 'г. Астана',
    '2': 'г. Алматы',
    '3': 'г. Шымкент',
    '4': 'Актобе Г.А.'
}

# УНИВЕРСАЛЬНЫЙ маппинг результатов визита
RESULT_MAPPING = {
    # Английский
    'Contact established - door opened': 'Контакт установлен',
    'No contact - no one opened the door': 'Неконтакт',
    'Household refused': 'Отказ',
    'Language barrier': 'Языковой барьер',
    'Invalid address - building demolished/under construction': 'Недоступный адрес',
    'Other (specify)': 'Другое',
    # Русский
    'Контакт установлен - дверь открыли': 'Контакт установлен',
    'Неконтакт - никто не открыл дверь': 'Неконтакт',
    'Отказ домохозяйства': 'Отказ',
    'Языковой барьер': 'Языковой барьер',
    'Недоступный адрес - здание снесено/в стройке': 'Недоступный адрес',
    'Другое (уточните)': 'Другое',
    # Казахский (если используется)
    'Байланыс орнатылды - есік ашылды': 'Контакт установлен',
    'Байланыс жоқ - ешкім есікті ашпады': 'Неконтакт',
    'Үй шаруашылығы бас тартты': 'Отказ'
}

# Готовность отвечать (все языки)
WILLINGNESS_YES = [
    'Yes, willing to answer now',  # English
    'Да, готов отвечать сейчас',  # Russian
    'Иә, қазір жауап беруге дайынмын'  # Kazakh (если есть)
]

# Согласие участвовать (все языки)
CONSENT_YES = [
    'Yes, I agree to participate',  # English
    'Да, согласен(на) принять участие',  # Russian
    'Да, соглашусь принять участие',  # Russian (альтернатива)
    'Иә, қатысуға келісемін'  # Kazakh (если есть)
]

# КВОТЫ
QUOTAS = {
    'г. Астана': {
        'total': 800,
        'employed': 608,
        'self_employed': 192,
        'peo_count': 32
    },
    'г. Алматы': {
        'total': 975,
        'employed': 741,
        'self_employed': 234,
        'peo_count': 39
    },
    'г. Шымкент': {
        'total': 700,
        'employed': 532,
        'self_employed': 168,
        'peo_count': 28
    },
    'Актобе Г.А.': {
        'total': 525,
        'employed': 399,
        'self_employed': 126,
        'peo_count': 21
    }
}

def fetch_kobo_data():
    """Загрузка данных из Kobo"""
    url = f"https://kf.kobotoolbox.org/api/v2/assets/{ASSET_ID}/data.json"
    headers = {'Authorization': f'Token {KOBO_API_TOKEN}'}
    
    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    
    data = response.json()
    return data.get('results', [])

def process_record(record):
    """Обработка одной записи"""
    # Правильные названия полей из Kobo
    city_raw = record.get('City:', '')
    city = CITY_MAPPING.get(city_raw, city_raw)  # Универсальный маппинг
    
    # Если город не распознан - попробуем найти частичное совпадение
    if city == city_raw and city_raw:
        for key, value in CITY_MAPPING.items():
            if key.lower() in city_raw.lower() or city_raw.lower() in key.lower():
                city = value
                break
    
    # Результат визита
    result_raw = record.get('Visit result:', '')
    result = RESULT_MAPPING.get(result_raw, result_raw if result_raw else 'Неизвестно')
    
    # Готовность и согласие (проверяем все языки)
    willingness = record.get('Is the respondent willing to answer?', '')
    consent = record.get('Are you willing to participate in this survey?', '')
    
    # Категория работника
    q08 = record.get('q08_survey2', '')
    
    # Другие поля
    interviewer = record.get('Interviewer full name:', '')
    peo = record.get('PEO number (electoral precinct)', '')
    
    # Определяем завершенность (проверка по всем языкам)
    is_completed = (
        willingness in WILLINGNESS_YES and
        consent in CONSENT_YES
    )
    
    # Определяем категорию
    if is_completed and str(q08).strip():
        q08_val = str(q08).strip()
        if q08_val == '1':
            category = 'employed'  # Наемный работник
        elif q08_val in ['2', '3', '4', '5']:
            category = 'self_employed'  # Самозанятый/ИП
        else:
            category = 'other'
    else:
        category = 'other'
    
    # Проверка на контакт (проверяем и английский и русский)
    is_contact = (
        'Contact established' in result or 
        'Контакт установлен' in result or
        'Байланыс орнатылды' in result
    )
    
    is_refusal = (
        'refused' in result.lower() or 
        'Отказ' in result or
        'бас тартты' in result
    )
    
    return {
        'city': city,
        'interviewer': interviewer,
        'peo': peo,
        'result': result,
        'category': category,
        'is_completed': is_completed,
        'is_contact': is_contact,
        'is_refusal': is_refusal
    }

def calculate_statistics(records):
    """Вычисление статистики"""
    processed = [process_record(r) for r in records]
    
    total_visits = len(processed)
    completed = sum(1 for r in processed if r['is_completed'])
    contacts = sum(1 for r in processed if r['is_contact'])
    refusals = sum(1 for r in processed if r['is_refusal'])
    
    # Метрики
    response_rate = round((completed / total_visits * 100) if total_visits > 0 else 0, 1)
    contact_rate = round((contacts / total_visits * 100) if total_visits > 0 else 0, 1)
    refusal_rate = round((refusals / contacts * 100) if contacts > 0 else 0, 1)
    
    # По городам
    city_stats = {}
    for city_name, quota_info in QUOTAS.items():
        city_records = [r for r in processed if r['city'] == city_name]
        city_completed = sum(1 for r in city_records if r['is_completed'])
        city_contacts = sum(1 for r in city_records if r['is_contact'])
        city_employed = sum(1 for r in city_records if r['is_completed'] and r['category'] == 'employed')
        city_self_employed = sum(1 for r in city_records if r['is_completed'] and r['category'] == 'self_employed')
        
        city_stats[city_name] = {
            'visits': len(city_records),
            'completed': city_completed,
            'employed': city_employed,
            'self_employed': city_self_employed,
            'quota_total': quota_info['total'],
            'quota_employed': quota_info['employed'],
            'quota_self_employed': quota_info['self_employed'],
            'peo_count': quota_info['peo_count'],
            'progress': round((city_completed / quota_info['total'] * 100) if quota_info['total'] > 0 else 0, 2),
            'contact_rate': round((city_contacts / len(city_records) * 100) if len(city_records) > 0 else 0, 1)
        }
    
    # По категориям (общее)
    employed = sum(1 for r in processed if r['category'] == 'employed')
    self_employed = sum(1 for r in processed if r['category'] == 'self_employed')
    
    total_quota = sum(q['total'] for q in QUOTAS.values())
    total_employed_quota = sum(q['employed'] for q in QUOTAS.values())
    total_self_employed_quota = sum(q['self_employed'] for q in QUOTAS.values())
    
    return {
        'overview': {
            'total_visits': total_visits,
            'completed': completed,
            'total_quota': total_quota,
            'quota_progress': round((completed / total_quota * 100) if total_quota > 0 else 0, 2),
            'response_rate': response_rate,
            'contact_rate': contact_rate,
            'refusal_rate': refusal_rate
        },
        'by_city': city_stats,
        'by_category': {
            'employed': {
                'completed': employed,
                'quota': total_employed_quota,
                'progress': round((employed / total_employed_quota * 100) if total_employed_quota > 0 else 0, 2)
            },
            'self_employed': {
                'completed': self_employed,
                'quota': total_self_employed_quota,
                'progress': round((self_employed / total_self_employed_quota * 100) if total_self_employed_quota > 0 else 0, 2)
            }
        },
        'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        try:
            # Загрузка данных
            records = fetch_kobo_data()
            
            # Вычисление статистики
            stats = calculate_statistics(records)
            
            # Ответ
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            
            self.wfile.write(json.dumps(stats, ensure_ascii=False).encode('utf-8'))
            
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            
            error_response = {
                'error': str(e),
                'message': 'Ошибка загрузки данных'
            }
            self.wfile.write(json.dumps(error_response, ensure_ascii=False).encode('utf-8'))
