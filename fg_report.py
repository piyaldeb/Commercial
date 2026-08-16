"""
FG Stock report -> Google Sheets, in the exact layout of the FG Store
dashboard's "Export to Excel" output (Fg_Stock.xlsx).

Source: operation.details.retrive_data_from_operation_details([]), the same call
the dashboard makes (captured in FG report.har, action 2591). The endpoint
returns 42 fields per OA; the Excel export keeps 33 of them, drops the internal
id/company/price fields, rounds money to 2dp and sorts by OA id.

Layout reproduced from Fg_Stock.xlsx, verified row-for-row against the HAR
captured 10 seconds after that export:
  row 1  - 33 headers, labels verbatim (including the source's "Recived" spelling)
  row 2  - Total row across the qty/value columns
  row 3+ - one row per OA, ascending oa_id

Four defects in the source export are deliberately NOT reproduced - see
EXPORT_FIXES below.

Target sheet: https://docs.google.com/spreadsheets/d/1YRLVLKbrwXIziBAmtAzgEfEPH64Ycwu57ZZJrDO6aks
"""

import os
import sys
import json
import time
import base64
import logging
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime, timezone, timedelta

import requests
from requests.exceptions import RequestException
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2 import service_account
from dotenv import load_dotenv

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
if sys.stderr.encoding and sys.stderr.encoding.lower() != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8")

load_dotenv()
logging.basicConfig(stream=sys.stdout, level=logging.INFO, format="%(message)s")
log = logging.getLogger()

# ========= CONFIG ==========
ODOO_URL = (os.getenv("ODOO_URL") or "").rstrip("/")
DB = os.getenv("ODOO_DB")
USERNAME = os.getenv("ODOO_USERNAME")
PASSWORD = os.getenv("ODOO_PASSWORD")

SHEET_KEY = "1YRLVLKbrwXIziBAmtAzgEfEPH64Ycwu57ZZJrDO6aks"
WORKSHEET = "FG Stock"

# Tabs written by the previous two-sheet version of this script, superseded by
# the single Fg_Stock layout. Dropped on each run so the sheet stays clean.
OBSOLETE_WORKSHEETS = ["FG Store", "FG Store Details"]

ALLOWED_COMPANY_IDS = [2, 1, 3]   # as the dashboard sends: Head Office, Zipper, Metal Trims
DHAKA = timezone(timedelta(hours=6))

EXPORT_FIXES = """\
  1. Total row alignment - the source export writes the totals two columns to the
     left of their headers (the Order QTY total lands under 'Item'). Here each
     total sits under its own column. 'Age' is left blank, as in the source.
  2. LC Number leading zeros - the export coerced them to numbers and lost the
     zeros on 46 of 443 rows ('0000019525120040' -> 19525120040). Written as text.
  3. Date cells - the export stamped every date with a spurious 06:00:20 time
     (a UTC+6 conversion artifact). Written as plain dates.
  4. Half-up rounding - matches the export's JS toFixed(2) on the 6 cells where
     Python's default banker's rounding would disagree (280.125 -> 280.13)."""

# label, API field, kind. Order and labels are verbatim from Fg_Stock.xlsx.
COLUMNS = [
    ("OA",              "oa_name",           "text"),
    ("Order Date",      "date_order",        "date"),
    ("ED Date",         "ed_date",           "date"),
    ("Closing Date",    "closing_date",      "date"),
    ("Sample",          "sample",            "text"),
    ("PI",              "pi",                "text"),
    ("Customer",        "customer_name",     "text"),
    ("Buyer",           "buyer_name",        "text"),
    ("Invoice No",      "invoice_number",    "text"),
    ("Invoice Date",    "invoice_date",      "date"),
    ("LC Number",       "lc_number",         "text"),
    ("LC Date",         "lc_date",           "date"),
    ("Sales Person",    "sales_person_name", "text"),
    ("Sales Team",      "sales_team",        "text"),
    ("Region",          "region_name",       "text"),
    ("DSM",             "core_leader_name",  "text"),
    ("Item",            "fg_categ_type",     "text"),
    ("Product",         "product_id",        "text"),
    ("Order QTY",       "order_qty",         "qty"),
    ("Order Value",     "order_value",       "value"),
    ("Recived QTY",     "received_qty",      "qty"),
    ("Recived Value",   "received_value",    "value"),
    ("Goods In Date",   "goods_in_date",     "date"),
    ("Delivered QTY",   "delivered_qty",     "qty"),
    ("Delivered Value", "delivered_value",   "value"),
    ("Delivered Date",  "delivery_date",     "date"),
    ("Pending QTY",     "pending_qty",       "qty"),
    ("Pending Value",   "pending_value",     "value"),
    ("Stock QTY",       "stock_qty",         "qty"),
    ("Stock Value",     "stock_value",       "value"),
    ("Age",             "days_passed",       "age"),
    ("Invoice QTY",     "invoice_qty",       "qty"),
    ("Invoice Value",   "invoice_value",     "value"),
]
LABELS = [label for label, _, _ in COLUMNS]

session = requests.Session()
USER_ID = None


# ========= VALUE COERCION ==========
def round2(v):
    """2dp, half away from zero - what the export's JS toFixed(2) does. Python's
    built-in round() uses banker's rounding and disagrees on exact .xx5 values."""
    if v is None or v == "":
        return ""
    return float(Decimal(float(v)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def clean_text(v):
    """Collapse whitespace runs and strip, as the export does. Source data has
    entries like 'MIRAS  FASHION  LTD' and 'KANIZ GARMENTS LIMITED\\n'."""
    if v is None or v is False:
        return ""
    return " ".join(str(v).split())


def clean_date(v):
    if not v:
        return ""
    return str(v)[:10]


# ========= RETRY ==========
def retry_request(method, url, max_retries=3, backoff=5, **kwargs):
    for attempt in range(1, max_retries + 1):
        try:
            r = method(url, **kwargs)
            r.raise_for_status()
            return r
        except RequestException as e:
            log.info(f"Attempt {attempt} failed: {e}")
            if attempt < max_retries:
                time.sleep(backoff)
            else:
                raise


# ========= ODOO ==========
def login():
    global USER_ID
    payload = {"jsonrpc": "2.0", "params": {"db": DB, "login": USERNAME, "password": PASSWORD}}
    r = retry_request(session.post, f"{ODOO_URL}/web/session/authenticate", json=payload, timeout=180)
    result = r.json().get("result") or {}
    if "uid" not in result:
        raise Exception("Odoo login failed")
    USER_ID = result["uid"]
    log.info(f"Logged in to Odoo (uid={USER_ID})")


def call_kw(model, method, args, timeout=900):
    payload = {
        "jsonrpc": "2.0", "method": "call",
        "params": {"model": model, "method": method, "args": args,
                   "kwargs": {"context": {"lang": "en_US", "tz": "Asia/Dhaka",
                                          "uid": USER_ID,
                                          "allowed_company_ids": ALLOWED_COMPANY_IDS}}},
    }
    r = retry_request(session.post, f"{ODOO_URL}/web/dataset/call_kw/{model}/{method}",
                      json=payload, timeout=timeout)
    body = r.json()
    if "error" in body:
        err = body["error"]
        raise Exception(f"{model}.{method} failed: {err.get('message')} :: "
                        f"{(err.get('data') or {}).get('message', '')}")
    return body.get("result")


def fetch_stock():
    """Rows sorted by ascending oa_id - the dashboard's JS iterates the datas
    object by its integer-like keys, which is the order the export preserves."""
    result = call_kw("operation.details", "retrive_data_from_operation_details", [[]]) or {}
    datas = result.get("datas") or {}
    rows = []
    for oa_id, group in datas.items():
        for rec in (group or []):
            rows.append((int(oa_id), rec))
    rows.sort(key=lambda t: t[0])
    log.info(f"Fetched {len(rows)} OA rows")
    return [rec for _, rec in rows]


# ========= TABLE ==========
def build_frame(records):
    """Header + Total + data, exactly the Fg_Stock.xlsx shape."""
    data = []
    for rec in records:
        row = {}
        for label, key, kind in COLUMNS:
            v = rec.get(key)
            if kind == "text":
                row[label] = clean_text(v)
            elif kind == "date":
                row[label] = clean_date(v)
            elif kind == "age":
                row[label] = int(v) if isinstance(v, (int, float)) else ""
            else:
                row[label] = round2(v)
        data.append(row)

    # Totals from the raw values, then rounded once - matching the export.
    total = {label: "" for label in LABELS}
    total["OA"] = "Total"
    for label, key, kind in COLUMNS:
        if kind in ("qty", "value"):
            total[label] = round2(sum(rec.get(key) or 0 for rec in records))

    return pd.DataFrame([total] + data, columns=LABELS)


# ========= GOOGLE SHEETS ==========
def get_gspread_client():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for sa_path in [os.path.join(script_dir, "service_account.json"), "service_account.json"]:
        if os.path.exists(sa_path):
            with open(sa_path, "r", encoding="utf-8") as f:
                log.info(f"Google service account: {json.load(f).get('client_email')}")
            return gspread.service_account(filename=sa_path)

    scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_dict = None
    creds_json = os.getenv("GOOGLE_SHEET_CRED_JSON")
    if creds_json:
        creds_dict = json.loads(creds_json)
    else:
        creds_raw = (os.getenv("GOOGLE_CREDS_BASE64") or "").strip()
        if creds_raw:
            try:
                creds_dict = json.loads(creds_raw)
            except json.JSONDecodeError:
                padded = creds_raw + "=" * (-len(creds_raw) % 4)
                creds_dict = json.loads(base64.b64decode(padded).decode("utf-8"))
    if creds_dict is None:
        raise Exception("No Google credentials found "
                        "(service_account.json / GOOGLE_SHEET_CRED_JSON / GOOGLE_CREDS_BASE64).")

    log.info(f"Google service account: {creds_dict.get('client_email')}")
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def drop_obsolete(sheet):
    existing = {ws.title: ws for ws in sheet.worksheets()}
    for title in OBSOLETE_WORKSHEETS:
        ws = existing.get(title)
        # Never delete the last remaining tab - Sheets rejects it anyway.
        if ws is not None and len(existing) > 1:
            sheet.del_worksheet(ws)
            existing.pop(title)
            log.info(f"Removed obsolete worksheet '{title}'")


def push(sheet, df):
    try:
        ws = sheet.worksheet(WORKSHEET)
    except gspread.WorksheetNotFound:
        ws = sheet.add_worksheet(title=WORKSHEET, rows=max(1000, len(df) + 50),
                                 cols=max(40, len(LABELS) + 2))
        log.info(f"Created worksheet '{WORKSHEET}'")

    need_rows, need_cols = max(len(df) + 50, 200), max(len(LABELS) + 2, 40)
    if ws.row_count < need_rows or ws.col_count < need_cols:
        ws.resize(rows=max(ws.row_count, need_rows), cols=max(ws.col_count, need_cols))

    ws.clear()
    sid = ws.id

    def fmt(idx, number_format, start_row=1):
        return {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": start_row,
                      "startColumnIndex": idx, "endColumnIndex": idx + 1},
            "cell": {"userEnteredFormat": {"numberFormat": number_format}},
            "fields": "userEnteredFormat.numberFormat"}}

    # clear() wipes values but not formatting; reset so last run's formats can't
    # reinterpret this run's values.
    sheet.batch_update({"requests": [
        {"updateCells": {"range": {"sheetId": sid}, "fields": "userEnteredFormat"}}]})

    # Text format must land BEFORE the values: USER_ENTERED honours an existing
    # TEXT format, and that is what keeps the leading zeros on LC numbers.
    text_reqs = [fmt(i, {"type": "TEXT"})
                 for i, (_, _, kind) in enumerate(COLUMNS) if kind == "text"]
    sheet.batch_update({"requests": text_reqs})

    set_with_dataframe(ws, df, resize=False)

    patterns = {"date": {"type": "DATE", "pattern": "yyyy-mm-dd"},
                "qty":  {"type": "NUMBER", "pattern": "#,##0.##"},
                "value": {"type": "NUMBER", "pattern": "#,##0.00"},
                "age":  {"type": "NUMBER", "pattern": "0"}}
    reqs = [fmt(i, patterns[kind])
            for i, (_, _, kind) in enumerate(COLUMNS) if kind in patterns]

    header_style = {"textFormat": {"bold": True},
                    "backgroundColor": {"red": 0.10, "green": 0.31, "blue": 0.58},
                    "horizontalAlignment": "CENTER"}
    header_style["textFormat"]["foregroundColor"] = {"red": 1, "green": 1, "blue": 1}
    reqs.append({"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
        "cell": {"userEnteredFormat": header_style},
        "fields": "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment)"}})
    reqs.append({"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": 2},
        "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.85, "green": 0.89, "blue": 0.96}}},
        "fields": "userEnteredFormat(textFormat,backgroundColor)"}})
    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 2}},
        "fields": "gridProperties.frozenRowCount"}})
    sheet.batch_update({"requests": reqs})

    log.info(f"Wrote {len(df) - 1} data rows + Total row x {len(LABELS)} cols to '{WORKSHEET}'")


# ========= MAIN ==========
def main():
    missing = [n for n, v in [("ODOO_URL", ODOO_URL), ("ODOO_DB", DB),
                              ("ODOO_USERNAME", USERNAME), ("ODOO_PASSWORD", PASSWORD)] if not v]
    if missing:
        raise SystemExit(f"Missing env vars: {', '.join(missing)}")

    login()
    df = build_frame(fetch_stock())

    totals = df.iloc[0]
    log.info("Totals: " + "  ".join(
        f"{lbl}={totals[lbl]:,.2f}" for lbl, _, kind in COLUMNS if kind in ("qty", "value")))

    if os.getenv("FG_REPORT_DRY_RUN"):
        out = os.getenv("FG_REPORT_DRY_RUN_PATH", "fg_stock_dryrun.csv")
        df.to_csv(out, index=False)
        log.info(f"DRY RUN: wrote {out} ({len(df) - 1} data rows), skipped Google Sheets")
        return

    client = get_gspread_client()
    sheet = client.open_by_key(SHEET_KEY)
    # Write first, drop second: the tab we just wrote guarantees the spreadsheet
    # still has a sheet left, so neither delete can hit the last-tab guard.
    push(sheet, df)
    drop_obsolete(sheet)
    log.info(f"Done at {datetime.now(DHAKA):%Y-%m-%d %H:%M} Asia/Dhaka")


if __name__ == "__main__":
    main()
