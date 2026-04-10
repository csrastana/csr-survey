"""
Serverless функция для получения статистики опроса
URL: /api/stats
С ПОДДЕРЖКОЙ ПАГИНАЦИИ - получает ВСЕ записи
"""

from http.server import BaseHTTPRequestHandler
import json
import re
import requests
from datetime import datetime, date

# Конфигурация
KOBO_API_TOKEN = '929c90ea6bbce9e24789c10b2eb9740e3352d859'
ASSET_ID = 'aCE5fencfcUpVhvCRdCoxc'

# ---- Фильтр по дате визита для Шымкента ----
# Исключаем анкеты по г. Шымкент с датой визита ДО этой даты (включительно).
# Т.е. учитываем только визиты 26 марта 2026 и позже.
SHYMKENT_MIN_DATE = date(2026, 3, 26)

# Возможные имена поля даты визита в группе group_xn8xb93.
# Код сначала попытается найти значение по этим именам,
# затем — автодетектом по любому ключу группы, значение которого парсится как дата.
VISIT_DATE_FIELDS = [
    'group_xn8xb93/date',        # подтверждено по коду download.py
    'group_xn8xb93/visit_date',
    'group_xn8xb93/date_visit',
    'group_xn8xb93/v_date',
]

DATE_RE = re.compile(r'(\d{4}-\d{2}-\d{2})')


def extract_visit_date(record):
    """
    Возвращает datetime.date даты визита или None.
    Сначала пробует явные имена, затем сканирует все ключи group_xn8xb93/,
    в крайнем случае берёт _submission_time.
    """
    # 1) Явные имена
    for key in VISIT_DATE_FIELDS:
        val = record.get(key)
        if val:
            m = DATE_RE.search(str(val))
            if m:
                try:
                    return datetime.strptime(m.group(1), '%Y-%m-%d').date()
                except ValueError:
                    pass

    # 2) Автодетект по группе
    for key, val in record.items():
        if not key.startswith('group_xn8xb93/') or val is None:
            continue
        m = DATE_RE.search(str(val))
        if m:
            try:
                return datetime.strptime(m.group(1), '%Y-%m-%d').date()
            except ValueError:
                continue

    # 3) Фоллбэк — время отправки на сервер
    sub = record.get('_submission_time')
    if sub:
        m = DATE_RE.search(str(sub))
        if m:
            try:
                return datetime.strptime(m.group(1), '%Y-%m-%d').date()
            except ValueError:
                pass
    return None


def get_validation_status(record):
    """
    Возвращает 'approved', 'not_approved' или 'no_status'.
    Kobo хранит _validation_status как dict с ключом 'uid'.
    """
    vs = record.get('_validation_status')
    if isinstance(vs, dict):
        uid = vs.get('uid', '')
        if uid == 'validation_status_approved':
            return 'approved'
        if uid == 'validation_status_not_approved':
            return 'not_approved'
    return 'no_status'

# Маппинг кодов городов
CITY_CODES = {
    '710000000': 'г. Астана',
    '750000000': 'г. Алматы',
    '790000000': 'г. Шымкент',
    '151000000': 'Актобе Г.А.'
}

# Маппинг кодов результатов
RESULT_CODES = {
    '1': 'Контакт установлен',
    '2': 'Неконтакт',
    '3': 'Недоступный адрес',
    '4': 'Языковой барьер',
    '5': 'Отказ',
    '6': 'Другое'
}

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
    """
    Загрузка ВСЕХ данных из Kobo с пагинацией
    Kobo API возвращает максимум 100 записей за раз - нужна пагинация!
    """
    headers = {'Authorization': f'Token {KOBO_API_TOKEN}'}
    
    all_results = []
    url = f"https://kf.kobotoolbox.org/api/v2/assets/{ASSET_ID}/data.json?limit=100"
    
    while url:
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        
        data = response.json()
        results = data.get('results', [])
        all_results.extend(results)
        
        # Следующая страница (если есть)
        url = data.get('next')
        
        # Безопасность: останавливаемся после 50 страниц (5000 записей)
        if len(all_results) > 5000:
            break
    
    return all_results

def process_record(record):
    """Обработка одной записи"""
    # Правильные названия полей из XLSForm
    city_code = record.get('city', '')
    city = CITY_CODES.get(str(city_code), f'Неизвестно ({city_code})')
    
    peo = record.get('group_xn8xb93/PEO', '')
    interviewer = record.get('group_xn8xb93/int_name', '')
    
    # Результат визита
    result_code = record.get('group_ip3jm92/result', '')
    result = RESULT_CODES.get(str(result_code), f'Неизвестно ({result_code})')
    
    # Готовность и согласие (коды!)
    willingness = record.get('willingness', '')
    consent = record.get('consent', '')
    
    # Категория работника
    q08 = record.get('q08_survey2', '')
    
    # Определяем завершенность (по кодам!)
    is_completed = (
        str(willingness) == '1' and  # '1' = Yes, willing to answer now
        str(consent) == '1'          # '1' = Yes, I agree to participate
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
    
    # Проверка на контакт
    is_contact = str(result_code) == '1'  # '1' = Contact established
    is_refusal = str(result_code) == '5'  # '5' = Household refusal

    # Статус валидации и дата визита
    validation = get_validation_status(record)
    visit_date = extract_visit_date(record)

    return {
        'city': city,
        'interviewer': interviewer,
        'peo': peo,
        'result': result,
        'category': category,
        'is_completed': is_completed,
        'is_contact': is_contact,
        'is_refusal': is_refusal,
        'validation': validation,
        'visit_date': visit_date,
    }

def calculate_statistics(records):
    """Вычисление статистики"""
    processed = [process_record(r) for r in records]

    # --- Фильтр по Шымкенту: исключаем визиты до SHYMKENT_MIN_DATE ---
    # Для прочих городов — ничего не трогаем.
    def keep(r):
        if r['city'] != 'г. Шымкент':
            return True
        vd = r['visit_date']
        if vd is None:
            # нет даты → не можем подтвердить, что это валидная анкета после 26 марта → исключаем
            return False
        return vd >= SHYMKENT_MIN_DATE

    processed = [r for r in processed if keep(r)]

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
        city_visits = len(city_records)
        city_completed = sum(1 for r in city_records if r['is_completed'])
        city_contacts = sum(1 for r in city_records if r['is_contact'])
        city_employed = sum(1 for r in city_records if r['is_completed'] and r['category'] == 'employed')
        city_self_employed = sum(1 for r in city_records if r['is_completed'] and r['category'] == 'self_employed')

        # Участие в опросе
        city_refusals = sum(1 for r in city_records if r['is_refusal'])
        city_not_agreed = city_visits - city_completed
        city_not_agreed_other = city_not_agreed - city_refusals  # неконтакты / барьер / другое / контакт без согласия
        agreement_rate = round((city_completed / city_visits * 100) if city_visits > 0 else 0, 1)

        # Валидация по городу (считаем по всем анкетам города)
        city_approved = sum(1 for r in city_records if r['validation'] == 'approved')
        city_not_approved = sum(1 for r in city_records if r['validation'] == 'not_approved')
        city_no_status = sum(1 for r in city_records if r['validation'] == 'no_status')

        city_stats[city_name] = {
            'visits': city_visits,
            'completed': city_completed,
            'employed': city_employed,
            'self_employed': city_self_employed,
            'quota_total': quota_info['total'],
            'quota_employed': quota_info['employed'],
            'quota_self_employed': quota_info['self_employed'],
            'peo_count': quota_info['peo_count'],
            'progress': round((city_completed / quota_info['total'] * 100) if quota_info['total'] > 0 else 0, 2),
            'contact_rate': round((city_contacts / city_visits * 100) if city_visits > 0 else 0, 1),
            'agreed': city_completed,
            'not_agreed': city_not_agreed,
            'refusals': city_refusals,
            'not_agreed_other': city_not_agreed_other,
            'agreement_rate': agreement_rate,
            'approved': city_approved,
            'not_approved': city_not_approved,
            'no_status': city_no_status,
        }
    
    # По категориям (общее)
    employed = sum(1 for r in processed if r['category'] == 'employed')
    self_employed = sum(1 for r in processed if r['category'] == 'self_employed')

    # Валидация (общая)
    total_approved = sum(1 for r in processed if r['validation'] == 'approved')
    total_not_approved = sum(1 for r in processed if r['validation'] == 'not_approved')
    total_no_status = sum(1 for r in processed if r['validation'] == 'no_status')

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
            'refusal_rate': refusal_rate,
            'approved': total_approved,
            'not_approved': total_not_approved,
            'no_status': total_no_status,
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
            # Загрузка ВСЕХ данных (с пагинацией)
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
