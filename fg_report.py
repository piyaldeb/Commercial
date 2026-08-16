"""
FG Store report → Google Sheets.

Replays exactly the two data calls the Odoo "FG Store" dashboard makes
(captured in FG report.har, action_id 2591 / taps_manufacturing.actions_fg_store_dashboard):

  1. operation.details.retrieve_fg_store_datas([1, 3], date_from, date_to)
       -> movement per OA + item: opening / received / delivery / disposal / closing
          (qty and value) plus LC status. The HAR captured 2026-08-01 -> 2026-08-16,
          i.e. month-to-date, which is what this script reproduces each run.

  2. operation.details.retrive_data_from_operation_details([])
       -> per-OA stock detail: buyer, customer, salesperson, team, region, core
          leader, product, order/received/delivered/pending/stock qty+value,
          invoice + LC references and ageing (days_passed).

Both payloads are written verbatim - no filtering, no re-aggregation. Only the
column headers are humanised; every value is exactly what the endpoint returned.

Target sheet: https://docs.google.com/spreadsheets/d/1YRLVLKbrwXIziBAmtAzgEfEPH64Ycwu57ZZJrDO6aks
"""

import os
import sys
import json
import time
import base64
import logging
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
STORE_WORKSHEET = "FG Store"
DETAIL_WORKSHEET = "FG Store Details"

# Exactly as in the HAR: the dashboard asks for Zipper (1) + Metal Trims (3),
# with Head Office (2) present in allowed_company_ids for record-rule purposes.
COMPANY_IDS = [1, 3]
ALLOWED_COMPANY_IDS = [2, 1, 3]

DHAKA = timezone(timedelta(hours=6))  # Asia/Dhaka, no DST

session = requests.Session()
USER_ID = None


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


def odoo_context():
    return {
        "lang": "en_US",
        "tz": "Asia/Dhaka",
        "uid": USER_ID,
        "allowed_company_ids": ALLOWED_COMPANY_IDS,
    }


def call_kw(model, method, args, timeout=900):
    payload = {
        "jsonrpc": "2.0", "method": "call",
        "params": {"model": model, "method": method, "args": args,
                   "kwargs": {"context": odoo_context()}},
    }
    r = retry_request(session.post, f"{ODOO_URL}/web/dataset/call_kw/{model}/{method}",
                      json=payload, timeout=timeout)
    body = r.json()
    if "error" in body:
        err = body["error"]
        raise Exception(f"{model}.{method} failed: {err.get('message')} :: "
                        f"{(err.get('data') or {}).get('message', '')}")
    return body.get("result")


def date_range():
    """Month-to-date in Asia/Dhaka, matching the range captured in the HAR
    (2026-08-01 -> 2026-08-16). Override with FG_DATE_FROM / FG_DATE_TO."""
    today = datetime.now(DHAKA).date()
    date_from = os.getenv("FG_DATE_FROM") or today.replace(day=1).isoformat()
    date_to = os.getenv("FG_DATE_TO") or today.isoformat()
    return date_from, date_to


# ========= DATASET 1: FG Store movement ==========
STORE_COLS = [
    ("oa",             "OA"),
    ("item",           "Item"),
    ("in_date",        "In Date"),
    ("opening_qty",    "Opening QTY"),
    ("opening_value",  "Opening Value"),
    ("received_qty",   "Received QTY"),
    ("received_value", "Received Value"),
    ("delivery_qty",   "Delivery QTY"),
    ("delivery_value", "Delivery Value"),
    ("disposal_qty",   "Disposal QTY"),
    ("disposal_value", "Disposal Value"),
    ("closing_qty",    "Closing QTY"),
    ("closing_value",  "Closing Value"),
    ("lc_status",      "LC Status"),
]
STORE_TEXT = {"oa", "item", "lc_status"}
STORE_DATE = {"in_date"}


def fetch_fg_store(date_from, date_to):
    rows = call_kw("operation.details", "retrieve_fg_store_datas",
                   [COMPANY_IDS, date_from, date_to]) or []
    log.info(f"FG Store movement {date_from} -> {date_to}: {len(rows)} rows")
    return rows


# ========= DATASET 2: per-OA stock detail ==========
DETAIL_COLS = [
    ("goods_in_date",     "Goods In Date"),
    ("buyer_id",          "Buyer ID"),
    ("buyer_name",        "Buyer"),
    ("company_id",        "Company ID"),
    ("company_name",      "Company"),
    ("customer_id",       "Customer ID"),
    ("customer_name",     "Customer"),
    ("fg_categ_type",     "FG Category"),
    ("oa_id",             "OA ID"),
    ("oa_name",           "OA"),
    ("date_order",        "Order Date"),
    ("ed_date",           "ED Date"),
    ("closing_date",      "Closing Date"),
    ("sample",            "Sample"),
    ("pi",                "PI"),
    ("invoice_number",    "Invoice Number"),
    ("invoice_date",      "Invoice Date"),
    ("lc_number",         "LC Number"),
    ("lc_date",           "LC Date"),
    ("delivery_date",     "Delivery Date"),
    ("sales_person_id",   "Salesperson ID"),
    ("sales_person_name", "Salesperson"),
    ("sales_team",        "Sales Team"),
    ("region_id",         "Region ID"),
    ("region_name",       "Region"),
    ("core_leader_id",    "Core Leader ID"),
    ("core_leader_name",  "Core Leader"),
    ("product_id",        "Product"),
    ("order_qty",         "Order QTY"),
    ("order_value",       "Order Value"),
    ("received_qty",      "Received QTY"),
    ("received_value",    "Received Value"),
    ("delivered_qty",     "Delivered QTY"),
    ("delivered_value",   "Delivered Value"),
    ("pending_qty",       "Pending QTY"),
    ("pending_value",     "Pending Value"),
    ("stock_qty",         "Stock QTY"),
    ("stock_value",       "Stock Value"),
    ("days_passed",       "Days Passed"),
    ("invoice_qty",       "Invoice QTY"),
    ("invoice_value",     "Invoice Value"),
    ("final_price",       "Final Price"),
]
# Identifier-ish columns that must stay verbatim strings. lc_number in particular
# carries leading zeros ("0000019525120040") that Sheets would otherwise eat.
DETAIL_TEXT = {"buyer_name", "company_name", "customer_name", "fg_categ_type",
               "oa_name", "sample", "pi", "invoice_number", "lc_number",
               "sales_person_name", "sales_team", "region_name",
               "core_leader_name", "product_id"}
DETAIL_DATE = {"goods_in_date", "date_order", "ed_date", "closing_date",
               "invoice_date", "lc_date", "delivery_date"}
DETAIL_INT = {"buyer_id", "company_id", "customer_id", "oa_id",
              "sales_person_id", "region_id", "core_leader_id", "days_passed"}


def fetch_detail():
    result = call_kw("operation.details", "retrive_data_from_operation_details", [[]]) or {}
    datas = result.get("datas") or {}
    rows = [row for group in datas.values() for row in (group or [])]
    log.info(f"FG Store details: {len(rows)} rows across {len(datas)} OAs")
    return rows


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
        creds_raw = os.getenv("GOOGLE_CREDS_BASE64")
        if creds_raw:
            raw = creds_raw.strip()
            try:
                creds_dict = json.loads(raw)
            except json.JSONDecodeError:
                padded = raw + "=" * (-len(raw) % 4)
                creds_dict = json.loads(base64.b64decode(padded).decode("utf-8"))

    if creds_dict is None:
        raise Exception("No Google credentials found "
                        "(service_account.json / GOOGLE_SHEET_CRED_JSON / GOOGLE_CREDS_BASE64).")

    # Printed so the first run tells you which address the target sheet must be shared with.
    log.info(f"Google service account: {creds_dict.get('client_email')}")
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def build_df(rows, colspec):
    """Frame with every column in the endpoint's own order, values untouched."""
    df = pd.DataFrame(rows, columns=[k for k, _ in colspec])
    df.columns = [label for _, label in colspec]
    return df


def push(sheet, title, df, colspec, text_keys, date_keys, int_keys=frozenset()):
    try:
        ws = sheet.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sheet.add_worksheet(title=title, rows=max(1000, len(df) + 50),
                                 cols=max(26, len(df.columns) + 2))
        log.info(f"Created worksheet '{title}'")

    need_rows = max(len(df) + 50, 100)
    need_cols = max(len(df.columns) + 2, 26)
    if ws.row_count < need_rows or ws.col_count < need_cols:
        ws.resize(rows=max(ws.row_count, need_rows), cols=max(ws.col_count, need_cols))

    ws.clear()

    sid = ws.id
    keys = [k for k, _ in colspec]

    def col_range(idx, from_data_row=True):
        return {"sheetId": sid, "startRowIndex": 1 if from_data_row else 0,
                "startColumnIndex": idx, "endColumnIndex": idx + 1}

    def fmt_req(idx, number_format):
        return {"repeatCell": {"range": col_range(idx),
                               "cell": {"userEnteredFormat": {"numberFormat": number_format}},
                               "fields": "userEnteredFormat.numberFormat"}}

    # clear() wipes values, not formatting - reset so a previous run's formats
    # can't reinterpret this run's values (e.g. numbers rendered as dates).
    sheet.batch_update({"requests": [
        {"updateCells": {"range": {"sheetId": sid}, "fields": "userEnteredFormat"}}
    ]})

    # Plain-text format must land BEFORE the values: USER_ENTERED input honours an
    # existing TEXT format, which is what preserves leading zeros in LC numbers.
    text_reqs = [fmt_req(i, {"type": "TEXT"})
                 for i, k in enumerate(keys) if k in text_keys]
    if text_reqs:
        sheet.batch_update({"requests": text_reqs})

    set_with_dataframe(ws, df, resize=False)

    # Numeric / date presentation for everything else.
    reqs = []
    for i, k in enumerate(keys):
        if k in text_keys:
            continue
        if k in date_keys:
            reqs.append(fmt_req(i, {"type": "DATE", "pattern": "yyyy-mm-dd"}))
        elif k in int_keys:
            reqs.append(fmt_req(i, {"type": "NUMBER", "pattern": "0"}))
        elif pd.api.types.is_numeric_dtype(df[df.columns[i]]):
            reqs.append(fmt_req(i, {"type": "NUMBER", "pattern": "#,##0.####"}))
    reqs.append({"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
        "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                       "backgroundColor": {"red": 0.85, "green": 0.89, "blue": 0.96}}},
        "fields": "userEnteredFormat(textFormat,backgroundColor)"}})
    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1}},
        "fields": "gridProperties.frozenRowCount"}})
    sheet.batch_update({"requests": reqs})

    log.info(f"Wrote {len(df)} rows x {len(df.columns)} cols to '{title}'")


def log_totals(df, label, cols):
    parts = [f"{c}={df[c].sum():,.2f}" for c in cols if c in df.columns]
    log.info(f"{label} totals: " + "  ".join(parts))


# ========= MAIN ==========
def main():
    missing = [n for n, v in [("ODOO_URL", ODOO_URL), ("ODOO_DB", DB),
                              ("ODOO_USERNAME", USERNAME), ("ODOO_PASSWORD", PASSWORD)] if not v]
    if missing:
        raise SystemExit(f"Missing env vars: {', '.join(missing)}")

    login()
    date_from, date_to = date_range()

    store_df = build_df(fetch_fg_store(date_from, date_to), STORE_COLS)
    detail_df = build_df(fetch_detail(), DETAIL_COLS)

    log_totals(store_df, "FG Store",
               ["Opening QTY", "Opening Value", "Received QTY", "Received Value",
                "Delivery QTY", "Delivery Value", "Closing QTY", "Closing Value"])
    log_totals(detail_df, "FG Store Details",
               ["Order QTY", "Order Value", "Received QTY", "Received Value",
                "Stock QTY", "Stock Value", "Pending QTY", "Pending Value"])

    if os.getenv("FG_REPORT_DRY_RUN"):
        store_df.to_csv("fg_store_dryrun.csv", index=False)
        detail_df.to_csv("fg_store_details_dryrun.csv", index=False)
        log.info("DRY RUN: wrote fg_store_dryrun.csv + fg_store_details_dryrun.csv, "
                 "skipped Google Sheets")
        return

    client = get_gspread_client()
    sheet = client.open_by_key(SHEET_KEY)

    push(sheet, STORE_WORKSHEET, store_df, STORE_COLS, STORE_TEXT, STORE_DATE)
    push(sheet, DETAIL_WORKSHEET, detail_df, DETAIL_COLS, DETAIL_TEXT, DETAIL_DATE, DETAIL_INT)

    log.info(f"Done at {datetime.now(DHAKA):%Y-%m-%d %H:%M} Asia/Dhaka "
             f"(FG Store range {date_from} -> {date_to})")


if __name__ == "__main__":
    main()
