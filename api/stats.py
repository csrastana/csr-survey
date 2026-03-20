"""
Serverless функция для получения статистики опроса
URL: /api/stats
"""

from http.server import BaseHTTPRequestHandler
import json
import requests
from datetime import datetime

# Конфигурация
KOBO_API_TOKEN = '929c90ea6bbce9e24789c10b2eb9740e3352d859'
ASSET_ID = 'aCE5fencfcUpVhvCRdCoxc'

# Маппинги
CITY_MAPPING = {
    '1': 'г. Астана',
    '2': 'г. Алматы', 
    '3': 'г. Шымкент',
    '4': 'Актобе Г.А.'
}

RESULT_MAPPING = {
    'Контакт установлен - дверь открыли': 'Контакт установлен',
    'Неконтакт - никто не открыл дверь': 'Неконтакт',
    'Отказ домохозяйства': 'Отказ',
    'Языковой барьер': 'Языковой барьер',
    'Недоступный адрес - здание снесено/в стройке': 'Недоступный адрес',
    'Другое (уточните)': 'Другое'
}

# ОБНОВЛЕННЫЕ КВОТЫ
# Астана = 32 ПЕО * 25 человек (19 наемных + 6 самозанятых)
# Алматы = 39 ПЕО * 25 человек (19 наемных + 6 самозанятых)
# Актобе = 21 ПЕО * 25 человек (19 наемных + 6 самозанятых)
# Шымкент = 28 ПЕО * 25 человек (19 наемных + 6 самозанятых)
QUOTAS = {
    'г. Астана': {
        'total': 800,        # 32 ПЕО * 25
        'employed': 608,     # 32 ПЕО * 19
        'self_employed': 192, # 32 ПЕО * 6
        'peo_count': 32
    },
    'г. Алматы': {
        'total': 975,        # 39 ПЕО * 25
        'employed': 741,     # 39 ПЕО * 19
        'self_employed': 234, # 39 ПЕО * 6
        'peo_count': 39
    },
    'г. Шымкент': {
        'total': 700,        # 28 ПЕО * 25
        'employed': 532,     # 28 ПЕО * 19
        'self_employed': 168, # 28 ПЕО * 6
        'peo_count': 28
    },
    'Актобе Г.А.': {
        'total': 525,        # 21 ПЕО * 25
        'employed': 399,     # 21 ПЕО * 19
        'self_employed': 126, # 21 ПЕО * 6
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
    # Извлекаем поля из групп Kobo
    result_text = record.get('group_ip3jm92/result', '')
    willingness = record.get('willingness', '')
    consent = record.get('consent', '')
    q08 = record.get('q08_survey2', '')
    city_code = record.get('city', '')
    int_name = record.get('group_xn8xb93/int_name', '')
    peo = record.get('group_xn8xb93/PEO', '')
    
    # Определяем город
    city = CITY_MAPPING.get(str(city_code), 'Неизвестно')
    
    # Определяем результат визита
    result = RESULT_MAPPING.get(result_text, result_text if result_text else 'Неизвестно')
    
    # Определяем категорию
    is_completed = (
        willingness == 'Да, готов отвечать сейчас' and
        consent in ['Да, согласен(на) принять участие', 'Да, соглашусь принять участие']
    )
    
    if is_completed and str(q08).strip():
        if str(q08).strip() == '1':
            category = 'employed'
        elif str(q08).strip() in ['2', '3', '4', '5']:
            category = 'self_employed'
        else:
            category = 'other'
    else:
        category = 'other'
    
    # Проверка на контакт
    is_contact = result in ['Контакт установлен']
    is_refusal = result == 'Отказ'
    
    return {
        'city': city,
        'interviewer': int_name,
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
