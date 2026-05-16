"""
add_from_url.py — добавление оборудования в библиотеку по ссылке на товар.

Использование:
    .venv/bin/python3 scripts/add_from_url.py "https://entero.ru/item/122375"
    .venv/bin/python3 scripts/add_from_url.py "https://..." --category "Кофемашины"
    .venv/bin/python3 scripts/add_from_url.py "https://..." --dry-run   # только вывод, без записи

Что делает:
    1. Скачивает страницу товара (requests + BeautifulSoup)
    2. Отправляет текст в GPT-4o-mini → получает JSON с характеристиками
    3. Добавляет строку в «Библиотека оборудования» в Google Sheets
"""

import argparse
import json
import os
import re
import sys
import time
from pathlib import Path
from typing import Optional

import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv
from openai import OpenAI
import googleapiclient.discovery
import google.oauth2.service_account

# ──────────────────────────────────────────
# Конфиг
# ──────────────────────────────────────────

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '..', '.env'))

SPREADSHEET_ID       = os.getenv('GOOGLE_SPREADSHEET_ID', '')
SERVICE_ACCOUNT_JSON = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON', '')
OPENAI_API_KEY       = os.getenv('OPENAI_API_KEY', '')

LIBRARY_SHEET   = 'Бибилиотека оборудования'  # с опечаткой как в оригинале
LIB_DATA_START  = 3   # данные начинаются со строки 3

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Категории из библиотеки — GPT выберет одну из них
KNOWN_CATEGORIES = [
    'Кофемашины', 'Кофемолки', 'Аксессуары эспрессо', 'Фильтр кофе',
    'Лёд', 'Холодильное оборудование', 'Блендеры', 'Соковыжималки',
    'Горячая вода', 'Водоподготовка', 'Спешиалти напитки', 'Посудомойка',
    'Аксессуары бара', 'Сантехника', 'Касса', 'Кухонное оборудование',
]

# ──────────────────────────────────────────
# 1. Загрузка страницы
# ──────────────────────────────────────────

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept-Language': 'ru-RU,ru;q=0.9',
}

def fetch_page_text(url: str) -> str:
    """Скачивает страницу с retry/backoff (3 попытки) и возвращает очищенный текст."""
    resp = None
    for attempt in range(3):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=20)
            if resp.status_code in (429, 503) and attempt < 2:
                wait = 3 * (2 ** attempt)  # 3s, 6s
                print(f'  ⏳ HTTP {resp.status_code}, повтор через {wait}с…')
                time.sleep(wait)
                resp = None
                continue
            resp.raise_for_status()
            break
        except requests.RequestException as exc:
            if attempt == 2:
                raise
            wait = 2 * (2 ** attempt)  # 2s, 4s
            print(f'  ⏳ Ошибка сети: {exc}. Повтор через {wait}с…')
            time.sleep(wait)
    if resp is None:
        raise requests.RequestException('Не удалось загрузить страницу после 3 попыток')
    soup = BeautifulSoup(resp.text, 'html.parser')

    # Удаляем шум: скрипты, стили, навигацию, футер
    for tag in soup(['script', 'style', 'nav', 'footer', 'header', 'aside', 'noscript']):
        tag.decompose()

    text = soup.get_text(separator='\n', strip=True)

    # Убираем пустые строки и ограничиваем размер (GPT лимит токенов)
    lines = [l for l in text.splitlines() if l.strip()]
    text  = '\n'.join(lines)
    return text[:12000]  # ~3000 токенов — достаточно для карточки товара


# ──────────────────────────────────────────
# 2. Извлечение данных через GPT
# ──────────────────────────────────────────

SYSTEM_PROMPT = f"""
Ты помогаешь заполнять базу данных барного оборудования.
Тебе дают текст страницы товара. Тебе нужно извлечь характеристики и вернуть JSON.

Верни ТОЛЬКО валидный JSON без markdown, без пояснений:
{{
  "name":       "Общее название (например: Кофемашина, Кофемолка, Холодильник)",
  "producer":   "Производитель и точное название модели (например: La Marzocco Linea PB 2Gr)",
  "dimensions": "габариты самого прибора БЕЗ упаковки строго в порядке: Ширина × Глубина × Высота в мм (например: 560x580x430). Если на сайте порядок другой — переставь правильно. Если несколько вариантов — бери рабочие (меньшие) размеры. Если не найдено — пустая строка.",
  "kw":         "Потребляемая мощность в кВт (например: 3.5). Только число. Если не найдено — пустая строка.",
  "water":      "Подключение к воде: 'Да' или 'Нет' или пустая строка если неизвестно.",
  "price":      "Цена в рублях (только число, без символов). Если не найдено — пустая строка.",
  "category":   "Одна из категорий: {', '.join(KNOWN_CATEGORIES)}. Выбери наиболее подходящую."
}}

Правила:
- producer должен содержать бренд + модель, как написано на сайте
- dimensions: если указаны ВхШхГ или другой порядок — переставь в Ш×Г×В
- kw: если указаны Вт — переведи в кВт (делить на 1000)
- price: только число, без пробелов и символов валюты
""".strip()


def extract_with_gpt(page_text: str, url: str) -> dict:
    """Отправляет текст страницы в GPT, получает структурированные данные."""
    client = OpenAI(api_key=OPENAI_API_KEY)

    user_msg = f"URL: {url}\n\nТекст страницы:\n{page_text}"

    response = client.chat.completions.create(
        model='gpt-4o-mini',
        messages=[
            {'role': 'system', 'content': SYSTEM_PROMPT},
            {'role': 'user',   'content': user_msg},
        ],
        temperature=0,
        max_tokens=500,
    )

    raw = response.choices[0].message.content.strip()

    # Убираем markdown-блоки если GPT всё же добавил
    raw = re.sub(r'^```[a-z]*\n?', '', raw)
    raw = re.sub(r'\n?```$', '', raw)

    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        print(f'❌ GPT вернул невалидный JSON:\n{raw}')
        raise e


# ──────────────────────────────────────────
# 3. Запись в Google Sheets
# ──────────────────────────────────────────

def get_sheets_service():
    creds = google.oauth2.service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_JSON, scopes=SCOPES
    )
    return googleapiclient.discovery.build('sheets', 'v4', credentials=creds)


def find_next_empty_row(service) -> int:
    """Находит первую пустую строку в библиотеке."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{LIBRARY_SHEET}'!A:A",
    ).execute()
    values = result.get('values', [])
    return max(len(values) + 1, LIB_DATA_START)


def check_duplicate(service, producer: str) -> Optional[int]:
    """Возвращает номер строки если такой producer уже есть в библиотеке."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{LIBRARY_SHEET}'!B:B",
    ).execute()
    values = result.get('values', [])
    norm = lambda s: s.strip().lower()
    for i, row in enumerate(values):
        if row and norm(row[0]) == norm(producer):
            return i + 1
    return None


def write_to_library(service, data: dict, row: int):
    """
    Записывает строку в библиотеку.
    Структура: A=Название, B=Производитель, C=Размеры, D=кВт, E=Вода, F=Цена, G=Категория
    """
    # Готовим строку: A..G (7 столбцов)
    row_data = [
        data.get('name',       ''),   # A
        data.get('producer',   ''),   # B
        data.get('dimensions', ''),   # C
        data.get('kw',         ''),   # D
        data.get('water',      ''),   # F→E
        data.get('price',      ''),   # G→F
        data.get('category',   ''),   # I→G
    ]

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{LIBRARY_SHEET}'!A{row}",
        valueInputOption='USER_ENTERED',
        body={'values': [row_data]},
    ).execute()


# ──────────────────────────────────────────
# 4. Главная функция
# ──────────────────────────────────────────

def process_one(url: str, category: Optional[str], dry_run: bool, service) -> bool:
    """Обрабатывает одну ссылку. Возвращает True если успешно."""
    print(f'\n🔗 Загружаю: {url}')
    try:
        page_text = fetch_page_text(url)
    except requests.RequestException as e:
        print(f'  ❌ Не удалось загрузить: {e}')
        return False

    print(f'  🤖 GPT ({len(page_text)} символов)...')
    try:
        data = extract_with_gpt(page_text, url)
    except Exception as e:
        print(f'  ❌ Ошибка GPT: {e}')
        return False

    if category:
        data['category'] = category

    print(f"  📋 {data.get('producer', '—')} | {data.get('category', '—')} | {data.get('price', '—')} ₽")

    if dry_run:
        print(f"     Размеры: {data.get('dimensions','—')}  кВт: {data.get('kw','—')}  Вода: {data.get('water','—')}")
        print('  ⚠️  dry-run: не записываю')
        return True

    dup_row = check_duplicate(service, data.get('producer', ''))
    if dup_row:
        print(f'  ⚠️  Уже есть в строке {dup_row} — пропускаю')
        return True

    next_row = find_next_empty_row(service)
    write_to_library(service, data, next_row)
    print(f'  ✅ Добавлено в строку {next_row}')
    return True


def main():
    parser = argparse.ArgumentParser(
        description='Добавить оборудование в библиотеку по ссылке на товар'
    )
    # Можно передать одну ссылку или файл со списком
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('url',    nargs='?', help='Ссылка на страницу товара')
    group.add_argument('--file', '-f',      help='Файл со списком ссылок (по одной на строку)')

    parser.add_argument('--category', default=None,
                        help='Принудительно задать категорию для всех ссылок')
    parser.add_argument('--dry-run', action='store_true',
                        help='Только вывести данные, не записывать в таблицу')
    args = parser.parse_args()

    # Собираем список URL
    if args.file:
        path = args.file if os.path.isabs(args.file) else os.path.join(
            os.path.dirname(__file__), '..', args.file)
        with open(path, encoding='utf-8') as f:
            urls = [
                line.strip() for line in f
                if line.strip() and not line.startswith('#')
            ]
        if not urls:
            print('⚠️  Файл пуст или все строки закомментированы')
            sys.exit(0)
        print(f'📄 Найдено {len(urls)} ссылок в {args.file}')
    else:
        urls = [args.url]

    # Валидация обязательных переменных окружения
    missing = [name for name, val in [
        ('GOOGLE_SPREADSHEET_ID',      SPREADSHEET_ID),
        ('GOOGLE_SERVICE_ACCOUNT_JSON', SERVICE_ACCOUNT_JSON),
        ('OPENAI_API_KEY',              OPENAI_API_KEY),
    ] if not val]
    if missing:
        print(f'\u274c Не заданы переменные в .env: {", ".join(missing)}', file=sys.stderr)
        sys.exit(1)
    if not args.dry_run and not Path(SERVICE_ACCOUNT_JSON).exists():
        print(f'\u274c Файл сервисного аккаунта не найден: {SERVICE_ACCOUNT_JSON}', file=sys.stderr)
        sys.exit(1)

    service = None if args.dry_run else get_sheets_service()

    ok = fail = 0
    for url in urls:
        if process_one(url, args.category, args.dry_run, service):
            ok += 1
        else:
            fail += 1

    print(f'\n{"─"*40}')
    print(f'Итого: ✅ {ok} добавлено  ❌ {fail} ошибок')
    if ok > 0 and not args.dry_run:
        print('💡 Нажмите «🔄 Обновить выпадающие списки» в таблице')


if __name__ == '__main__':
    main()
