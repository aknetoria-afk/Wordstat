"""
Telegram бот для сбора статистики Wordstat.

Требования:
  pip install python-telegram-bot python-dotenv google-auth google-auth-httplib2 google-api-python-client

Запуск:
  python bot.py
"""

import os
import json
import http.client
import ssl
import time
import logging
from datetime import date, timedelta
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
import asyncio
from concurrent.futures import ThreadPoolExecutor
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    ContextTypes, ConversationHandler,
)

load_dotenv()

# ── Настройки ────────────────────────────────────────────────────────────────

TELEGRAM_TOKEN   = os.getenv("TELEGRAM_TOKEN")
YANDEX_API_KEY   = os.getenv("YANDEX_API_KEY")
YANDEX_FOLDER_ID = os.getenv("YANDEX_FOLDER_ID")
SPREADSHEET_ID   = os.getenv("SPREADSHEET_ID", "198_LrEC4b04EuguRXT3MmGSx1E4ENQeQHgezC1YFby8")
CREDENTIALS_FILE = "wordstat-google.json"
SPREADSHEET_URL  = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"

SHEETS = [
    "Айрис бренд",
    "Дабл бренд",
    "Смородина бренд",
    "Север бренд",
    "Спрос на недвижимость",
]
DEFAULT_SHEET = "Спрос на недвижимость"

MONTHS = [
    ("Январь",  1), ("Февраль", 2), ("Март",     3),
    ("Апрель",  4), ("Май",     5), ("Июнь",     6),
    ("Июль",    7), ("Август",  8), ("Сентябрь", 9),
    ("Октябрь", 10), ("Ноябрь", 11), ("Декабрь", 12),
]

YEAR = 2026

# Состояния диалога
SELECT_PROJECT, SELECT_PERIOD, SELECT_WEEK, SELECT_MONTH = range(4)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Google Sheets ─────────────────────────────────────────────────────────────

def get_sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds)


def get_sheet_id(service, sheet_name):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == sheet_name:
            return s["properties"]["sheetId"]
    return None


def read_keywords(service, sheet_name):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!1:1",
    ).execute()
    row = result.get("values", [[]])[0]
    return [cell for cell in row[1:] if cell and cell != "ИТОГО"]


def read_existing_data(service, sheet_name, keywords):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A2:ZZ",
    ).execute()
    rows = result.get("values", [])
    existing = {}
    for row in rows:
        if not row or not row[0]:
            continue
        label = row[0]
        if any(v for v in row[1:]):
            kw_dict = {}
            for i, kw in enumerate(keywords):
                val = int(row[i + 1]) if i + 1 < len(row) and row[i + 1] else 0
                kw_dict[kw] = val
            existing[label] = kw_dict
    return existing


def write_week_to_sheets(service, sheet_name, row_index, label, kw_dict, keywords):
    values = [kw_dict.get(kw, 0) for kw in keywords]
    itogo = sum(values)
    row = [label] + values + [itogo]
    wait = 5
    for attempt in range(5):
        try:
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"'{sheet_name}'!A{row_index}",
                valueInputOption="RAW",
                body={"values": [row]},
            ).execute()
            return
        except Exception as e:
            logger.warning(f"Ошибка записи: {e}, повтор через {wait} сек...")
            time.sleep(wait)
            wait *= 2


def get_weeks_for_month(year, month):
    """Возвращает все недели которые попадают в данный месяц."""
    weeks = []
    first_day = date(year, month, 1)
    if month == 12:
        last_day = date(year, 12, 31)
    else:
        last_day = date(year, month + 1, 1) - timedelta(days=1)

    # Начинаем с первого понедельника <= первого числа месяца
    current = first_day - timedelta(days=first_day.weekday())

    while current <= last_day:
        week_end = current + timedelta(days=6)
        # Включаем неделю если она пересекается с месяцем
        if week_end >= first_day and current <= last_day:
            actual_end = min(week_end, last_day)
            label = f"{current.strftime('%d.%m.%Y')}-{(current + timedelta(days=6)).strftime('%d.%m.%Y')}"
            weeks.append(label)
        current += timedelta(days=7)
    return weeks


def get_available_weeks():
    """Возвращает все недели от 05.01.2026 до сегодня."""
    weeks = []
    current = date(2026, 1, 5)
    today = date.today()
    while current <= today:
        week_end = current + timedelta(days=6)
        label = f"{current.strftime('%d.%m.%Y')}-{week_end.strftime('%d.%m.%Y')}"
        weeks.append(label)
        current += timedelta(days=7)
    return weeks


# ── Wordstat API ──────────────────────────────────────────────────────────────

def fetch_keyword_week(keyword, date_from, date_to):
    payload = {
        "phrase": keyword,
        "period": "PERIOD_WEEKLY",
        "fromDate": date_from,
        "toDate": date_to,
        "folderId": YANDEX_FOLDER_ID,
        "regions": ["11176"],
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {
        "Authorization": f"Api-key {YANDEX_API_KEY}",
        "Content-Type": "application/json; charset=utf-8",
    }
    wait = 5
    for attempt in range(5):
        ctx = ssl.create_default_context()
        conn = http.client.HTTPSConnection("searchapi.api.cloud.yandex.net", timeout=30, context=ctx)
        conn.request("POST", "/v2/wordstat/dynamics", body=body, headers=headers)
        resp = conn.getresponse()
        status = resp.status
        raw = resp.read()
        conn.close()
        if status != 429:
            break
        time.sleep(wait)
        wait *= 2

    if status in (500, 400):
        return 0
    if status >= 400:
        raise RuntimeError(f"HTTP {status}: {raw.decode('utf-8', errors='replace')}")
    data = json.loads(raw.decode("utf-8"))
    return sum(int(r.get("count", 0)) for r in data.get("results", []))


def collect_week(service, sheet_name, week_label):
    """Собирает данные за одну неделю и записывает в таблицу."""
    keywords = read_keywords(service, sheet_name)

    # Определяем номер строки
    all_weeks = get_available_weeks()
    try:
        row_index = all_weeks.index(week_label) + 2
    except ValueError:
        row_index = len(all_weeks) + 2

    # Парсим даты из метки
    parts = week_label.split("-")
    d_from = date(int(parts[0].split(".")[2]), int(parts[0].split(".")[1]), int(parts[0].split(".")[0]))
    d_to   = date(int(parts[1].split(".")[2]), int(parts[1].split(".")[1]), int(parts[1].split(".")[0]))
    date_from = d_from.strftime("%Y-%m-%dT00:00:00Z")
    date_to   = d_to.strftime("%Y-%m-%dT00:00:00Z")

    kw_dict = {}
    for kw in keywords:
        count = fetch_keyword_week(kw, date_from, date_to)
        kw_dict[kw] = count
        time.sleep(37)

    write_week_to_sheets(service, sheet_name, row_index, week_label, kw_dict, keywords)
    return kw_dict


# ── Хэндлеры бота ────────────────────────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [[InlineKeyboardButton(sheet, callback_data=f"project:{sheet}")] for sheet in SHEETS]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(
        "👋 Привет! Я бот для мониторинга поисковых запросов.\n\nВыберите проект:",
        reply_markup=reply_markup,
    )
    return SELECT_PROJECT


async def select_project(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    sheet_name = query.data.replace("project:", "")
    context.user_data["sheet_name"] = sheet_name

    keyboard = [
        [InlineKeyboardButton("📅 Неделя", callback_data="period:week")],
        [InlineKeyboardButton("📆 Месяц",  callback_data="period:month")],
    ]
    await query.edit_message_text(
        f"Проект: *{sheet_name}*\n\nВыберите период:",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return SELECT_PERIOD


async def select_period(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    period = query.data.replace("period:", "")
    context.user_data["period"] = period

    if period == "week":
        weeks = get_available_weeks()[-8:]  # последние 8 недель
        keyboard = [[InlineKeyboardButton(w, callback_data=f"week:{w}")] for w in reversed(weeks)]
        await query.edit_message_text(
            "Выберите неделю:",
            reply_markup=InlineKeyboardMarkup(keyboard),
        )
        return SELECT_WEEK
    else:
        keyboard = [
            [InlineKeyboardButton(name, callback_data=f"month:{num}")]
            for name, num in MONTHS
        ]
        await query.edit_message_text(
            "Выберите месяц:",
            reply_markup=InlineKeyboardMarkup(keyboard),
        )
        return SELECT_MONTH


async def select_week(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    week_label = query.data.replace("week:", "")
    sheet_name = context.user_data["sheet_name"]

    await query.edit_message_text(f"🔍 Проверяю данные за *{week_label}*...", parse_mode="Markdown")

    service = get_sheets_service()
    keywords = read_keywords(service, sheet_name)
    existing = read_existing_data(service, sheet_name, keywords)

    if week_label in existing:
        await query.edit_message_text(
            f"✅ Данные за *{week_label}* уже собраны!\n\n"
            f"📊 [Открыть таблицу]({SPREADSHEET_URL})",
            parse_mode="Markdown",
        )
    else:
        await query.edit_message_text(
            f"⏳ Данных за *{week_label}* нет. Начинаю сбор...\n"
            f"Это займёт несколько минут.",
            parse_mode="Markdown",
        )
        try:
            collect_week(service, sheet_name, week_label)
            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text=f"✅ Данные за *{week_label}* собраны!\n\n"
                     f"📊 [Открыть таблицу]({SPREADSHEET_URL})",
                parse_mode="Markdown",
            )
        except Exception as e:
            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text=f"❌ Ошибка при сборе данных: {e}",
            )

    return ConversationHandler.END


async def select_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    month_num  = int(query.data.replace("month:", ""))
    month_name = next(name for name, num in MONTHS if num == month_num)
    sheet_name = context.user_data["sheet_name"]

    await query.edit_message_text(
        f"🔍 Проверяю данные за *{month_name} {YEAR}*...",
        parse_mode="Markdown",
    )

    service  = get_sheets_service()
    keywords = read_keywords(service, sheet_name)
    existing = read_existing_data(service, sheet_name, keywords)
    weeks    = get_weeks_for_month(YEAR, month_num)

    missing = [w for w in weeks if w not in existing]

    if not missing:
        await query.edit_message_text(
            f"✅ Данные за *{month_name} {YEAR}* уже собраны!\n\n"
            f"📊 [Открыть таблицу]({SPREADSHEET_URL})",
            parse_mode="Markdown",
        )
    else:
        await query.edit_message_text(
            f"⏳ Не хватает {len(missing)} из {len(weeks)} недель за *{month_name} {YEAR}*.\n"
            f"Начинаю сбор... Это займёт ~{len(missing) * len(keywords) // 2 + 1} мин.",
            parse_mode="Markdown",
        )
        errors = []
        for week_label in missing:
            try:
                collect_week(service, sheet_name, week_label)
            except Exception as e:
                errors.append(f"{week_label}: {e}")

        if errors:
            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text=f"⚠️ Сбор завершён с ошибками:\n" + "\n".join(errors),
            )
        else:
            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text=f"✅ Данные за *{month_name} {YEAR}* собраны!\n\n"
                     f"📊 [Открыть таблицу]({SPREADSHEET_URL})",
                parse_mode="Markdown",
            )

    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Отменено. Напишите /start чтобы начать заново.")
    return ConversationHandler.END


# ── Запуск ────────────────────────────────────────────────────────────────────

async def main():

    app = Application.builder().token(TELEGRAM_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            SELECT_PROJECT: [CallbackQueryHandler(select_project, pattern="^project:")],
            SELECT_PERIOD:  [CallbackQueryHandler(select_period,  pattern="^period:")],
            SELECT_WEEK:    [CallbackQueryHandler(select_week,    pattern="^week:")],
            SELECT_MONTH:   [CallbackQueryHandler(select_month,   pattern="^month:")],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(conv_handler)

    logger.info("Бот запущен...")
    async with app:
        await app.start()
        await app.updater.start_polling(allowed_updates=Update.ALL_TYPES)
        await asyncio.Event().wait()


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
