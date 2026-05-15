"""
Скрипт для сбора статистики из Яндекс Wordstat (через Yandex Search API)
и записи данных + графиков в Google Таблицу.

Требования:
  pip install google-auth google-auth-httplib2 google-api-python-client

Использование:
  python wordstat_fill.py --apikey ВАШ_API_КЛЮЧ --folder ВАШ_FOLDER_ID
  python wordstat_fill.py --iam ВАШ_IAM_ТОКЕН --folder ВАШ_FOLDER_ID
"""

import argparse
import time
import json
import http.client
import ssl
from datetime import date, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── Настройки ────────────────────────────────────────────────────────────────

SPREADSHEET_ID   = "198_LrEC4b04EuguRXT3MmGSx1E4ENQeQHgezC1YFby8"
CREDENTIALS_FILE = "wordstat-google.json"
SKIP_SHEETS      = {"Диаграммы"}  # вкладки которые не трогаем

def _generate_weeks(start, end):
    weeks = []
    current = start
    while current <= end:
        week_end = min(current + timedelta(days=6), end)
        label = f"{current.strftime('%d.%m.%Y')}-{week_end.strftime('%d.%m.%Y')}"
        weeks.append((
            current.strftime("%Y-%m-%dT00:00:00Z"),
            week_end.strftime("%Y-%m-%dT00:00:00Z"),
            label,
        ))
        current += timedelta(days=7)
    return weeks

WEEKS = _generate_weeks(date(2026, 1, 5), date(2026, 4, 26))

# ── Google Sheets ─────────────────────────────────────────────────────────────

def get_service():
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds)


def get_all_sheets(service):
    """Возвращает список (title, sheetId) для всех рабочих вкладок."""
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = []
    for s in meta["sheets"]:
        title = s["properties"]["title"]
        if title not in SKIP_SHEETS:
            sheets.append((title, s["properties"]["sheetId"]))
    return sheets


def read_keywords(service, sheet_name):
    """Читает ключевики из строки 1 (столбец B и далее), без ИТОГО."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!1:1",
    ).execute()
    row = result.get("values", [[]])[0]
    return [cell for cell in row[1:] if cell and cell != "ИТОГО"]


def ensure_itogo_header(service, sheet_name, keywords):
    """Расширяет таблицу если нужно и добавляет заголовок ИТОГО."""
    # Сначала узнаём sheetId и текущие размеры
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_id = None
    current_cols = 0
    for s in meta["sheets"]:
        if s["properties"]["title"] == sheet_name:
            sheet_id = s["properties"]["sheetId"]
            current_cols = s["properties"]["gridProperties"]["columnCount"]
            break

    needed_cols = len(keywords) + 2  # A + ключевики + ИТОГО

    # Расширяем таблицу если нужно
    if current_cols < needed_cols:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_id,
                        "gridProperties": {"columnCount": needed_cols + 5},
                    },
                    "fields": "gridProperties.columnCount",
                }
            }]}
        ).execute()

    # Проверяем есть ли уже ИТОГО
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!1:1",
    ).execute()
    row = result.get("values", [[]])[0]
    itogo_col = len(keywords) + 2
    if len(row) >= itogo_col and row[itogo_col - 1] == "ИТОГО":
        return

    col_letter = col_num_to_letter(itogo_col)
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!{col_letter}1",
        valueInputOption="RAW",
        body={"values": [["ИТОГО"]]},
    ).execute()


def col_num_to_letter(n):
    """Конвертирует номер столбца (1-based) в букву (A, B, ..., Z, AA, ...)."""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def read_existing_data(service, sheet_name, keywords):
    """Читает уже заполненные строки. Возвращает dict {label: {kw: val}}."""
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
    """Записывает одну неделю + ИТОГО в указанную строку. Повторяет при ошибке сети."""
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
            print(f"  ⚠ Ошибка записи: {e}, повтор через {wait} сек... (попытка {attempt+1}/5)")
            time.sleep(wait)
            wait *= 2
    raise RuntimeError("Не удалось записать данные в Google Sheets после 5 попыток")


def ensure_chart_sheet(service):
    """Создаёт лист Диаграммы если его нет, возвращает sheetId."""
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta["sheets"]:
        if s["properties"]["title"] == "Диаграммы":
            return s["properties"]["sheetId"]
    body = {"requests": [{"addSheet": {"properties": {"title": "Диаграммы"}}}]}
    resp = service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body
    ).execute()
    return resp["replies"][0]["addSheet"]["properties"]["sheetId"]


def delete_existing_charts(service, chart_sheet_id):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for sheet in meta.get("sheets", []):
        if sheet["properties"]["sheetId"] == chart_sheet_id:
            chart_ids = [c["chartId"] for c in sheet.get("charts", [])]
            if chart_ids:
                requests = [{"deleteEmbeddedObject": {"objectId": cid}} for cid in chart_ids]
                service.spreadsheets().batchUpdate(
                    spreadsheetId=SPREADSHEET_ID, body={"requests": requests}
                ).execute()
            break


def add_chart_for_sheet(service, data_sheet_id, chart_sheet_id, sheet_name, keywords, weekly_data, anchor_row):
    """Добавляет линейный график для одной вкладки."""
    n_weeks  = len(weekly_data)
    n_kw     = len(keywords)

    line_series = []
    for col_idx in range(n_kw):
        line_series.append({
            "series": {
                "sourceRange": {
                    "sources": [{
                        "sheetId": data_sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": n_weeks + 1,
                        "startColumnIndex": col_idx + 1,
                        "endColumnIndex": col_idx + 2,
                    }]
                }
            },
            "targetAxis": "LEFT_AXIS",
            "dataLabel": {
                "type": "DATA",
                "placement": "ABOVE",
                "textFormat": {"fontSize": 8},
            },
        })

    request = {
        "addChart": {
            "chart": {
                "spec": {
                    "title": f"Динамика запросов — {sheet_name}",
                    "basicChart": {
                        "chartType": "LINE",
                        "legendPosition": "RIGHT_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Неделя"},
                            {
                                "position": "LEFT_AXIS",
                                "title": "Количество показов",
                                "viewWindowOptions": {
                                    "viewWindowMode": "EXPLICIT",
                                    "viewWindowMin": 0,
                                },
                            },
                        ],
                        "domains": [{
                            "domain": {
                                "sourceRange": {
                                    "sources": [{
                                        "sheetId": data_sheet_id,
                                        "startRowIndex": 1,
                                        "endRowIndex": n_weeks + 1,
                                        "startColumnIndex": 0,
                                        "endColumnIndex": 1,
                                    }]
                                }
                            }
                        }],
                        "series": line_series,
                        "headerCount": 1,
                    },
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": chart_sheet_id,
                            "rowIndex": anchor_row,
                            "columnIndex": 0,
                        },
                        "widthPixels": 900,
                        "heightPixels": 400,
                    }
                },
            }
        }
    }
    return request


def add_all_charts(service, all_sheets_data):
    """Создаёт по одному графику на каждую вкладку на листе Диаграммы."""
    chart_sheet_id = ensure_chart_sheet(service)
    delete_existing_charts(service, chart_sheet_id)

    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_ids = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}

    requests = []
    anchor_row = 0
    for sheet_name, keywords, weekly_data in all_sheets_data:
        if not keywords or not weekly_data:
            continue
        data_sheet_id = sheet_ids[sheet_name]
        req = add_chart_for_sheet(
            service, data_sheet_id, chart_sheet_id,
            sheet_name, keywords, weekly_data, anchor_row
        )
        requests.append(req)
        anchor_row += 22  # отступ между графиками

    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": requests},
        ).execute()
        print(f"  ✓ Добавлено {len(requests)} диаграмм на лист 'Диаграммы'")


# ── Wordstat API ──────────────────────────────────────────────────────────────

def fetch_keyword_week(auth_header, folder_id, keyword, date_from, date_to):
    payload = {
        "phrase": keyword,
        "period": "PERIOD_WEEKLY",
        "fromDate": date_from,
        "toDate": date_to,
        "folderId": folder_id,
        "regions": ["11176"],  # Тюменская область
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    headers = {
        "Authorization": auth_header,
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
        print(f"  Лимит запросов, ждём {wait} сек... (попытка {attempt + 1}/5)")
        time.sleep(wait)
        wait *= 2

    if status == 500:
        print(f"  ⚠ Данные недоступны (HTTP 500), пропускаем")
        return 0
    if status >= 400:
        raise RuntimeError(f"HTTP {status}: {raw.decode('utf-8', errors='replace')}")

    data = json.loads(raw.decode("utf-8"))
    total = sum(int(r.get("count", 0)) for r in data.get("results", []))
    return total


def fetch_all_keywords(auth_header, folder_id, keywords, date_from, date_to):
    result = {}
    for i, kw in enumerate(keywords, start=1):
        print(f"    [{i}/{len(keywords)}] {kw}")
        count = fetch_keyword_week(auth_header, folder_id, kw, date_from, date_to)
        result[kw] = count
        time.sleep(37)  # лимит 100 запросов/час
    return result


# ── Основной поток ────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Yandex Wordstat → Google Sheets (все вкладки)")
    group = parser.add_mutually_exclusive_group(required=False)
    group.add_argument("--apikey", help="API-ключ Яндекс")
    group.add_argument("--iam",    help="IAM-токен Яндекс")
    parser.add_argument("--folder", required=False, default="", help="folderId Yandex Cloud")
    args = parser.parse_args()

    print("Подключаемся к Google Sheets...")
    service = get_service()

    sheets = get_all_sheets(service)
    print(f"Найдено вкладок: {[s[0] for s in sheets]}\n")

    if not args.folder:
        print("Ошибка: укажите --folder для сбора данных")
        return

    auth_header = f"Api-key {args.apikey}" if args.apikey else f"Bearer {args.iam}"

    all_sheets_data = []

    for sheet_name, _ in sheets:
        print(f"\n{'='*50}")
        print(f"Вкладка: {sheet_name}")
        print(f"{'='*50}")

        keywords = read_keywords(service, sheet_name)
        if not keywords:
            print("  Нет ключевиков, пропускаем")
            continue

        print(f"Ключевиков: {len(keywords)} → {keywords}")

        ensure_itogo_header(service, sheet_name, keywords)
        existing = read_existing_data(service, sheet_name, keywords)
        print(f"Уже заполнено: {len(existing)} из {len(WEEKS)} недель\n")

        weekly_data = []
        for idx, (d_from, d_to, label) in enumerate(WEEKS, start=1):
            if label in existing:
                print(f"  [{idx}/{len(WEEKS)}] {label} — уже есть, пропускаем")
                weekly_data.append((label, existing[label]))
                continue

            print(f"  [{idx}/{len(WEEKS)}] {label}")
            kw_dict = fetch_all_keywords(auth_header, args.folder, keywords, d_from, d_to)
            weekly_data.append((label, kw_dict))

            row_index = idx + 1
            write_week_to_sheets(service, sheet_name, row_index, label, kw_dict, keywords)
            print(f"  ✓ Сохранено\n")

        all_sheets_data.append((sheet_name, keywords, weekly_data))

    print(f"\n✅ Готово! https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")


if __name__ == "__main__":
    main()
