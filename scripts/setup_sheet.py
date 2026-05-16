#!/usr/bin/env python3
"""
setup_sheet.py — первичная настройка таблицы «Барная эргономика».

Что делает:
  1. Читает «Библиотека оборудования»
  2. Добавляет столбец «Категория» (колонка I), если его нет
  3. Авто-заполняет категории по ключевым словам в названии
  4. Устанавливает выпадающий список в столбце B «Оборудование»

Запуск:
  python3 scripts/setup_sheet.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")

SPREADSHEET_ID = os.getenv("GOOGLE_SPREADSHEET_ID", "1XSVheSQkxr8H87Fm-c7JPwz472SBLWd6p72GH4gYALM")
SA_JSON        = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")

EQUIPMENT_SHEET = "Оборудование"
LIBRARY_SHEET   = "Бибилиотека оборудования"  # с опечаткой в оригинале
MAX_EQ_ROWS     = 200  # макс. число строк в вкладке Оборудование (для data validation)

# ──────────────────────────────────────────
# Правила категорий (порядок важен — от конкретного к общему)
# ──────────────────────────────────────────
CATEGORY_RULES: list[tuple[str, str]] = [
    # Кофе эспрессо
    ("кофемашина",              "Кофемашины"),
    # Кофемолки
    ("кофемолка для эспрессо",  "Кофемолки"),
    ("кофемолка для фильтра",   "Кофемолки"),
    ("кофемолка",               "Кофемолки"),
    # Аксессуары эспрессо
    ("темпер",                  "Аксессуары эспрессо"),
    ("дозатор молока",          "Аксессуары эспрессо"),
    # Фильтр кофе
    ("фильтр машина",           "Фильтр кофе"),
    ("фильтр-машина",           "Фильтр кофе"),
    ("термос для фильтр",       "Фильтр кофе"),
    ("термосы для фильтр",      "Фильтр кофе"),
    # Нитро / спешиалти
    ("нитро",                   "Спешиалти напитки"),
    # Лёд
    ("ледогенератор",           "Лёд"),
    ("льдогенератор",           "Лёд"),
    # Холодильное
    ("витрина холодильная",     "Холодильное оборудование"),
    ("шкаф холодильный",        "Холодильное оборудование"),
    ("стол холодильный",        "Холодильное оборудование"),
    ("холодильник",             "Холодильное оборудование"),
    ("морозильник",             "Холодильное оборудование"),
    # Горячая вода
    ("бойлер",                  "Горячая вода"),
    ("водонагреватель",         "Горячая вода"),
    # Водоподготовка
    ("система водоподготовки",  "Водоподготовка"),
    ("водоподготовка",          "Водоподготовка"),
    # Блендеры
    ("блендер",                 "Блендеры"),
    # Соки
    ("соковыжималка",           "Соковыжималки"),
    # Посудомойка
    ("посудомоечная",           "Посудомойка"),
    ("стаканомоечная",          "Посудомойка"),
    # Кухня
    ("пресс-гриль",             "Кухонное оборудование"),
    ("гриль",                   "Кухонное оборудование"),
    # Аксессуары
    ("нок-бокс",                "Аксессуары бара"),
    ("ринзер",                  "Аксессуары бара"),
    ("диспенсер для стаканов",  "Аксессуары бара"),
    ("диспенсер",               "Аксессуары бара"),
    ("отверстие для мусора",    "Аксессуары бара"),
    ("мусорное ведро",          "Аксессуары бара"),
    # Сантехника
    ("раковина",                "Сантехника"),
    # Касса
    ("кассовая",                "Касса"),
]


def assign_category(name: str) -> str:
    n = name.lower().strip()
    for keyword, cat in CATEGORY_RULES:
        if keyword in n:
            return cat
    return "Прочее"


def get_service():
    creds = Credentials.from_service_account_file(
        SA_JSON,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds)


def main() -> None:
    if not SA_JSON:
        print("ERROR: GOOGLE_SERVICE_ACCOUNT_JSON не задан в .env", file=sys.stderr)
        sys.exit(1)
    if not Path(SA_JSON).exists():
        print(f"ERROR: Файл SA JSON не найден: {SA_JSON}", file=sys.stderr)
        sys.exit(1)

    svc = get_service()
    ss = svc.spreadsheets()

    # ── 1. Читаем библиотеку ──────────────────────────────────
    print("1. Читаем библиотеку...")
    result = ss.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{LIBRARY_SHEET}'!A2:G200",
        valueRenderOption="FORMATTED_VALUE",
    ).execute()
    rows = result.get("values", [])
    rows = [(r + [""] * 7)[:7] for r in rows]   # дополняем до 7 ячеек

    data_rows = [r for r in rows[1:] if any(r)]
    print(f"   Строк данных: {len(data_rows)}")

    # ── 2. Добавляем/обновляем категории ─────────────────────
    print("2. Заполняем категории...")

    # Заголовок (строка 2 таблицы = rows[0])
    if rows:
        rows[0][6] = "Категория"

    updated = 0
    for row in rows[1:]:
        if not row[0].strip():
            continue
        if not row[6].strip():
            row[6] = assign_category(row[0])
            updated += 1

    # Вывод для проверки
    print(f"   Присвоено категорий: {updated}")
    cats: dict[str, list[str]] = {}
    for row in rows[1:]:
        if row[0].strip():
            cats.setdefault(row[6], []).append(row[0])
    for cat, items in sorted(cats.items()):
        print(f"   [{cat}] {len(items)} позиций: {', '.join(items[:3])}{'...' if len(items) > 3 else ''}")

    # ── 3. Записываем библиотеку обратно ─────────────────────
    print("3. Записываем категории в таблицу...")
    end_row = 2 + len(rows) - 1
    ss.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{LIBRARY_SHEET}'!A2:G{end_row}",
        valueInputOption="RAW",
        body={"values": rows},
    ).execute()
    print(f"   Записано {len(rows)} строк (включая заголовок)")

    # ── 4. Data validation на B2:B50 в Оборудовании ──────────
    print("4. Устанавливаем выпадающие списки (B=Категория) в «Оборудование» (B2:B50)...")

    # Уникальные категории из столбца I библиотеки (index 8)
    cats_seen: set[str] = set()
    categories: list[str] = []
    for row in rows[1:]:
        cat = row[6].strip() if len(row) > 6 else ""
        if cat and cat not in cats_seen:
            categories.append(cat)
            cats_seen.add(cat)
    categories.sort()
    print(f"   Категорий в библиотеке: {len(categories)}: {categories}")

    # sheetId вкладки Оборудование
    meta = ss.get(spreadsheetId=SPREADSHEET_ID).execute()
    equipment_sheet_id = None
    for s in meta["sheets"]:
        if s["properties"]["title"] == EQUIPMENT_SHEET:
            equipment_sheet_id = s["properties"]["sheetId"]
            break

    if equipment_sheet_id is None:
        print(f"   ⚠ Вкладка «{EQUIPMENT_SHEET}» не найдена", file=sys.stderr)
        return

    body = {
        "requests": [{
            "setDataValidation": {
                "range": {
                    "sheetId": equipment_sheet_id,
                    "startRowIndex": 1,          # B2
                    "endRowIndex": MAX_EQ_ROWS,  # B{MAX_EQ_ROWS}
                    "startColumnIndex": 1,
                    "endColumnIndex": 2,
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [{"userEnteredValue": n} for n in categories],
                    },
                    "showCustomUi": True,
                    "strict": False,
                    "inputMessage": "Выберите из библиотеки или введите вручную",
                },
            },
        }]
    }
    ss.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
    print(f"   ✅ Выпадающий список категорий на B2:B50 установлен ({len(categories)} категорий)")

    # ── Итог ─────────────────────────────────────────────────
    print()
    print("✅ Настройка завершена!")
    print(f"   Таблица: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
    print()
    print("СЛЕДУЮЩИЙ ШАГ — добавьте Apps Script в таблицу:")
    print("  1. Откройте таблицу → Расширения → Apps Script")
    print("  2. Вставьте содержимое apps-script/Code.gs в редактор (файл Code.gs)")
    print("  3. Создайте файл Dialog.html → вставьте содержимое apps-script/Dialog.html")
    print("  4. Сохраните → запустите функцию setupTriggers() один раз")
    print("  5. После этого в таблице появится меню «🔧 Оборудование»")


if __name__ == "__main__":
    main()
