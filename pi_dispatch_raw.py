"""
Commercial pipeline raw data -> Google Sheets, tab "Raw".

One long/tidy table, one row per source record, so the numbers can be pivoted in
the sheet (rows = Metric, columns = Date, filters = Month / Sales Team / Sales
Person / Customer) without any further shaping.

SCOPE: the current calendar month, companies 1 (Zipper) and 3 (Metal Trims)
combined - the same allowed_company_ids the dashboards send.

METRICS AND THEIR SOURCES. Each is the call captured in the matching .har, and
the first four were verified against taps.pi.dispatch.summary - the model behind
the FG Store dashboard's stage tiles, whose stages are exactly pi_issued /
oa_released / lc_received / packed / dispatched:

  PI issue           sale.order, sales_type='sale', state='sale', date_order in
                     the month.  (Sales Pi and OA.har)
                     Verified: matches the summary's pi_issued qty and value on
                     every day of the month, both companies.
  OA Released        sale.order, sales_type='oa', state='sale', date_order in
                     the month.  (Sales Pi and OA.har)
                     Verified: Metal Trims qty/value match oa_released exactly.
  LC Received        combine.invoice.line, parent_state='posted', invoice_date in
                     the month.  (Combined Invoice.har)  The lines rather than the
                     invoice header: an invoice covers both companies at once and
                     only the lines carry company_id and the salesperson.
                     Verified: row count, qty AND value match lc_received exactly
                     for both companies.
  Dispatch           operation.details, next_operation='Delivery',
                     state not in (cancel, closed), action_date in the month.
                     (Fg Delivery.har)  Value = qty * final_price.
                     Verified: row count, qty AND value match dispatched exactly
                     for both companies.

  FG Store Balance   operation.details.retrive_data_from_operation_details - the
                     FG Store dashboard's own call.  (Fg Store.har)
  Production Pending manufacturing.order, oa_total_balance > 0, oa_id set,
                     state not in (closed, cancel, hold).  (running order.har)
                     Value = balance_qty * final_price.

The last two are BALANCES, not monthly flows: they are what is outstanding right
now, including stock and orders older than this month, so they are NOT filtered
to the month. Each such row is stamped with the run date so it lands on today's
column in a date pivot; its own document date is kept in "Doc Date".

operation.details and manufacturing.order carry no salesperson, so Sales Person
(and Sales Team / Region where blank) is filled in from the linked sale.order via
oa_id.

Target sheet: https://docs.google.com/spreadsheets/d/1PI5KF_WpOIEb3zzqXuC2JNX4ntk6Pf5WkBg28eEM_kk
"""

import os
import sys
import json
import time
import base64
import logging
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime, timezone, timedelta, date

import requests
from requests.exceptions import RequestException
import pandas as pd
import gspread
from gspread.exceptions import APIError
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

SHEET_KEY = "1PI5KF_WpOIEb3zzqXuC2JNX4ntk6Pf5WkBg28eEM_kk"
WORKSHEET = "Raw"

COMPANY_IDS = [1, 3]
COMPANY_NAMES = {1: "Zipper", 3: "Metal Trims"}
DHAKA = timezone(timedelta(hours=6))

# Metric labels, in the order the user's pivot template lists them.
M_PI = "PI issue"
M_OA = "OA Released"
M_LC = "LC Received"
M_DISPATCH = "Dispatch"
M_FG = "FG Store Balance"
M_PROD = "Production Pending"
METRIC_ORDER = [M_PI, M_OA, M_LC, M_DISPATCH, M_FG, M_PROD]

# label, kind. "id" columns are written as text so LC / invoice numbers keep
# their leading zeros (set_with_dataframe writes USER_ENTERED).
COLUMNS = [
    ("Metric",       "text"),
    ("Date",         "date"),
    ("Month",        "id"),
    ("Company",      "text"),
    ("Sales Team",   "text"),
    ("Sales Person", "text"),
    ("Region",       "text"),
    ("Customer",     "text"),
    ("Buyer",        "text"),
    ("Doc No",       "id"),
    ("OA",           "id"),
    ("LC No",        "id"),
    ("Item",         "text"),
    ("Doc Date",     "date"),
    ("QTY",          "qty"),
    ("Value",        "value"),
    ("Status",       "text"),
]
LABELS = [label for label, _ in COLUMNS]
KIND = dict(COLUMNS)

session = requests.Session()
USER_ID = None


# ========= VALUE COERCION ==========
def round2(v):
    """2dp, half away from zero - what the dashboards' JS toFixed(2) does."""
    if v is None or v == "" or v is False:
        return 0.0
    return float(Decimal(float(v)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def clean_text(v):
    """Collapse whitespace runs and strip; the source has 'MIRAS  FASHION  LTD'
    and trailing newlines."""
    if v is None or v is False:
        return ""
    return " ".join(str(v).split())


def clean_date(v):
    if not v:
        return ""
    return str(v)[:10]


def m2o_name(v):
    """Odoo many2one comes back as [id, display_name] (or False)."""
    if not v:
        return ""
    if isinstance(v, (list, tuple)):
        return clean_text(v[1] if len(v) > 1 else "")
    return clean_text(v)


def m2o_id(v):
    if not v:
        return None
    if isinstance(v, (list, tuple)):
        return v[0]
    return v


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


def sheets_call(fn, *args, **kwargs):
    for attempt in range(1, 6):
        try:
            return fn(*args, **kwargs)
        except APIError as e:
            code = (e.response.status_code if getattr(e, "response", None) is not None else None)
            if code in (429, 500, 503) and attempt < 5:
                wait = 5 * attempt
                log.info(f"Sheets API {code}, retrying in {wait}s")
                time.sleep(wait)
                continue
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


def call_kw(model, method, args=None, timeout=900, **kwargs):
    payload = {
        "jsonrpc": "2.0", "method": "call",
        "params": {"model": model, "method": method, "args": args or [],
                   "kwargs": {**kwargs,
                              "context": {"lang": "en_US", "tz": "Asia/Dhaka",
                                          "uid": USER_ID,
                                          "allowed_company_ids": COMPANY_IDS}}},
    }
    r = retry_request(session.post, f"{ODOO_URL}/web/dataset/call_kw/{model}/{method}",
                      json=payload, timeout=timeout)
    body = r.json()
    if "error" in body:
        err = body["error"]
        raise Exception(f"{model}.{method} failed: {err.get('message')} :: "
                        f"{(err.get('data') or {}).get('message', '')}")
    return body.get("result")


def search_read(model, domain, fields, **kwargs):
    """limit=0 is 'all' - these result sets run to a few thousand rows."""
    recs = call_kw(model, "search_read", [domain, fields], limit=0, **kwargs) or []
    log.info(f"  {model}: {len(recs)} rows")
    return recs


# ========= MONTH WINDOW ==========
def month_window(today):
    start = today.replace(day=1)
    end = (start + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    return start, end


# ========= SOURCES ==========
SALE_FIELDS = ["name", "date_order", "company_id", "team_id", "user_id", "region_id",
               "partner_id", "buyer_name", "total_product_qty", "amount_total",
               "lc_number", "state", "sales_type", "order_ref"]


def fetch_sale_orders(sales_type, start, end):
    """PI issue / OA Released. date_order is the dashboard's stage date."""
    domain = [["sales_type", "=", sales_type],
              ["state", "=", "sale"],
              ["company_id", "in", COMPANY_IDS],
              ["date_order", ">=", f"{start} 00:00:00"],
              ["date_order", "<=", f"{end} 23:59:59"]]
    return search_read("sale.order", domain, SALE_FIELDS)


def fetch_invoice_lines(start, end):
    """LC Received, at line level - the header has no company_id and its
    sales_person many2many is empty in practice."""
    domain = [["parent_state", "=", "posted"],
              ["company_id", "in", COMPANY_IDS],
              ["invoice_date", ">=", str(start)],
              ["invoice_date", "<=", str(end)]]
    return search_read("combine.invoice.line", domain,
                       ["invoice_id", "invoice_date", "delivery_date", "company_id",
                        "team_id", "sales_person", "customer_id", "buyer_id",
                        "fg_categ_type", "sale_order", "pi_number",
                        "quantity", "price_subtotal", "parent_state"])


def fetch_invoice_headers(invoice_ids):
    """Invoice number and LC number live on the header."""
    invoice_ids = sorted({i for i in invoice_ids if i})
    lookup = {}
    for i in range(0, len(invoice_ids), 500):
        chunk = invoice_ids[i:i + 500]
        for rec in call_kw("combine.invoice", "read",
                           [chunk, ["name", "lc_no", "lc_date"]]) or []:
            lookup[rec["id"]] = rec
    log.info(f"  combine.invoice headers: {len(lookup)}")
    return lookup


def fetch_deliveries(start, end):
    """Dispatch. Value is qty * final_price - final_price is the unit price."""
    domain = [["next_operation", "=", "Delivery"],
              ["state", "not in", ["cancel", "closed"]],
              ["company_id", "in", COMPANY_IDS],
              ["action_date", ">=", f"{start} 00:00:00"],
              ["action_date", "<=", f"{end} 23:59:59"]]
    return search_read("operation.details", domain,
                       ["oa_id", "company_id", "team_id", "partner_id", "buyer_id",
                        "fg_categ_type", "delivery_code", "action_date", "date_order",
                        "qty", "final_price", "state"])


def fetch_fg_store():
    """FG Store Balance - the dashboard's own call, keyed by OA id."""
    result = call_kw("operation.details", "retrive_data_from_operation_details", [[]]) or {}
    datas = result.get("datas") or {}
    rows = []
    for oa_id, group in datas.items():
        for rec in (group or []):
            rows.append((int(oa_id), rec))
    rows.sort(key=lambda t: t[0])
    rows = [rec for _, rec in rows if rec.get("company_id") in COMPANY_IDS]
    log.info(f"  operation.details (FG store): {len(rows)} rows")
    return rows


def fetch_running_orders():
    """Production Pending - the running-order list, balance still to produce."""
    domain = ["&", "&", ["oa_total_balance", ">", 0], ["oa_id", "!=", False],
              ["state", "not in", ["closed", "cancel", "hold"]]]
    recs = search_read("manufacturing.order", domain,
                       ["oa_id", "company_id", "partner_id", "buyer_name", "team_id",
                        "date_order", "fg_categ_type", "product_template_id",
                        "product_uom_qty", "done_qty", "balance_qty", "final_price",
                        "state"])
    return [r for r in recs if m2o_id(r.get("company_id")) in COMPANY_IDS]


def fetch_oa_lookup(oa_ids):
    """operation.details and manufacturing.order carry no salesperson; their
    oa_id points at sale.order, which does."""
    oa_ids = sorted({i for i in oa_ids if i})
    lookup = {}
    for i in range(0, len(oa_ids), 500):
        chunk = oa_ids[i:i + 500]
        for rec in call_kw("sale.order", "read",
                           [chunk, ["name", "user_id", "team_id", "region_id",
                                    "partner_id", "buyer_name", "lc_number"]]) or []:
            lookup[rec["id"]] = rec
    log.info(f"  sale.order lookup: {len(lookup)} OAs")
    return lookup


# ========= ROW BUILDERS ==========
def blank_row():
    return {label: "" for label in LABELS}


def stamp(row, metric, day):
    row["Metric"] = metric
    row["Date"] = clean_date(day)
    row["Month"] = clean_date(day)[:7]
    return row


def rows_sale_orders(recs, metric):
    out = []
    for r in recs:
        row = blank_row()
        stamp(row, metric, r.get("date_order"))
        row["Company"] = m2o_name(r.get("company_id"))
        row["Sales Team"] = m2o_name(r.get("team_id"))
        row["Sales Person"] = m2o_name(r.get("user_id"))
        row["Region"] = m2o_name(r.get("region_id"))
        row["Customer"] = m2o_name(r.get("partner_id"))
        row["Buyer"] = m2o_name(r.get("buyer_name"))
        row["Doc No"] = clean_text(r.get("name"))
        row["OA"] = clean_text(r.get("name")) if metric == M_OA else m2o_name(r.get("order_ref"))
        row["LC No"] = clean_text(r.get("lc_number"))
        row["Doc Date"] = clean_date(r.get("date_order"))
        row["QTY"] = round2(r.get("total_product_qty"))
        row["Value"] = round2(r.get("amount_total"))
        row["Status"] = clean_text(r.get("state"))
        out.append(row)
    return out


def rows_invoice_lines(recs, headers, oa_lookup):
    out = []
    for r in recs:
        head = headers.get(m2o_id(r.get("invoice_id"))) or {}
        oa = oa_lookup.get(m2o_id(r.get("sale_order"))) or {}
        row = blank_row()
        stamp(row, M_LC, r.get("invoice_date"))
        row["Company"] = m2o_name(r.get("company_id"))
        row["Sales Team"] = m2o_name(r.get("team_id"))
        row["Sales Person"] = m2o_name(r.get("sales_person"))
        row["Customer"] = m2o_name(r.get("customer_id"))
        row["Buyer"] = m2o_name(r.get("buyer_id"))
        row["Region"] = m2o_name(oa.get("region_id"))
        row["Doc No"] = clean_text(head.get("name")) or m2o_name(r.get("invoice_id"))
        row["OA"] = m2o_name(r.get("sale_order"))
        row["LC No"] = clean_text(head.get("lc_no"))
        row["Item"] = clean_text(r.get("fg_categ_type"))
        row["Doc Date"] = clean_date(r.get("invoice_date"))
        row["QTY"] = round2(r.get("quantity"))
        row["Value"] = round2(r.get("price_subtotal"))
        row["Status"] = clean_text(r.get("parent_state"))
        out.append(row)
    return out


def rows_deliveries(recs, oa_lookup):
    out = []
    for r in recs:
        oa = oa_lookup.get(m2o_id(r.get("oa_id"))) or {}
        row = blank_row()
        stamp(row, M_DISPATCH, r.get("action_date"))
        row["Company"] = m2o_name(r.get("company_id"))
        row["Sales Team"] = m2o_name(r.get("team_id")) or m2o_name(oa.get("team_id"))
        row["Sales Person"] = m2o_name(oa.get("user_id"))
        row["Region"] = m2o_name(oa.get("region_id"))
        row["Customer"] = m2o_name(r.get("partner_id"))
        row["Buyer"] = m2o_name(r.get("buyer_id"))
        row["Doc No"] = clean_text(r.get("delivery_code"))
        row["OA"] = m2o_name(r.get("oa_id"))
        row["LC No"] = clean_text(oa.get("lc_number"))
        row["Item"] = clean_text(r.get("fg_categ_type"))
        row["Doc Date"] = clean_date(r.get("date_order"))
        row["QTY"] = round2(r.get("qty"))
        row["Value"] = round2((r.get("qty") or 0) * (r.get("final_price") or 0))
        row["Status"] = clean_text(r.get("state"))
        out.append(row)
    return out


def rows_fg_store(recs, as_of):
    """Balance as of the run date - stock still sitting in the FG store."""
    out = []
    for r in recs:
        row = blank_row()
        stamp(row, M_FG, as_of)
        row["Company"] = clean_text(r.get("company_name")) or COMPANY_NAMES.get(r.get("company_id"), "")
        row["Sales Team"] = clean_text(r.get("sales_team"))
        row["Sales Person"] = clean_text(r.get("sales_person_name"))
        row["Region"] = clean_text(r.get("region_name"))
        row["Customer"] = clean_text(r.get("customer_name"))
        row["Buyer"] = clean_text(r.get("buyer_name"))
        row["Doc No"] = clean_text(r.get("invoice_number"))
        row["OA"] = clean_text(r.get("oa_name"))
        row["LC No"] = clean_text(r.get("lc_number"))
        row["Item"] = clean_text(r.get("fg_categ_type"))
        row["Doc Date"] = clean_date(r.get("goods_in_date"))
        row["QTY"] = round2(r.get("stock_qty"))
        row["Value"] = round2(r.get("stock_value"))
        out.append(row)
    return out


def rows_running_orders(recs, oa_lookup, as_of):
    """Balance as of the run date - order quantity still to be produced."""
    out = []
    for r in recs:
        oa = oa_lookup.get(m2o_id(r.get("oa_id"))) or {}
        row = blank_row()
        stamp(row, M_PROD, as_of)
        row["Company"] = m2o_name(r.get("company_id"))
        row["Sales Team"] = m2o_name(r.get("team_id")) or m2o_name(oa.get("team_id"))
        row["Sales Person"] = m2o_name(oa.get("user_id"))
        row["Region"] = m2o_name(oa.get("region_id"))
        row["Customer"] = m2o_name(r.get("partner_id"))
        row["Buyer"] = m2o_name(r.get("buyer_name"))
        row["OA"] = m2o_name(r.get("oa_id"))
        row["LC No"] = clean_text(oa.get("lc_number"))
        row["Item"] = clean_text(r.get("fg_categ_type")) or m2o_name(r.get("product_template_id"))
        row["Doc Date"] = clean_date(r.get("date_order"))
        row["QTY"] = round2(r.get("balance_qty"))
        row["Value"] = round2((r.get("balance_qty") or 0) * (r.get("final_price") or 0))
        row["Status"] = clean_text(r.get("state"))
        out.append(row)
    return out


def build(as_of):
    start, end = month_window(as_of)
    log.info(f"Month window {start} .. {end}, companies {COMPANY_IDS}")

    log.info("Fetching PI issue / OA Released ...")
    pis = fetch_sale_orders("sale", start, end)
    oas = fetch_sale_orders("oa", start, end)

    log.info("Fetching LC Received ...")
    invoice_lines = fetch_invoice_lines(start, end)
    headers = fetch_invoice_headers([m2o_id(r.get("invoice_id")) for r in invoice_lines])

    log.info("Fetching Dispatch ...")
    deliveries = fetch_deliveries(start, end)

    log.info("Fetching FG Store Balance ...")
    fg = fetch_fg_store()

    log.info("Fetching Production Pending ...")
    running = fetch_running_orders()

    oa_lookup = fetch_oa_lookup([m2o_id(r.get("oa_id")) for r in deliveries]
                                + [m2o_id(r.get("oa_id")) for r in running]
                                + [m2o_id(r.get("sale_order")) for r in invoice_lines])

    rows = (rows_sale_orders(pis, M_PI)
            + rows_sale_orders(oas, M_OA)
            + rows_invoice_lines(invoice_lines, headers, oa_lookup)
            + rows_deliveries(deliveries, oa_lookup)
            + rows_fg_store(fg, as_of)
            + rows_running_orders(running, oa_lookup, as_of))

    df = pd.DataFrame(rows, columns=LABELS)
    order = {m: i for i, m in enumerate(METRIC_ORDER)}
    df = (df.assign(_m=df["Metric"].map(order))
            .sort_values(["_m", "Date", "Company", "Customer"], kind="stable")
            .drop(columns="_m")
            .reset_index(drop=True))

    for metric in METRIC_ORDER:
        sel = df[df["Metric"] == metric]
        log.info(f"  {metric:20} {len(sel):6} rows  qty={sel['QTY'].sum():15,.2f}  "
                 f"value={sel['Value'].sum():13,.2f}")
    return df


# ========= GOOGLE SHEETS ==========
def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds_dict = None
    if os.path.exists("service_account.json"):
        with open("service_account.json", encoding="utf-8") as fh:
            creds_dict = json.load(fh)
    else:
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


NUMBER_PATTERNS = {"date": {"type": "DATE", "pattern": "yyyy-mm-dd"},
                   "qty": {"type": "NUMBER", "pattern": "#,##0.##"},
                   "value": {"type": "NUMBER", "pattern": "#,##0.00"}}

BLUE = {"red": 0.10, "green": 0.31, "blue": 0.58}


def col_range(sheet_id, index, rows):
    return {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": rows,
            "startColumnIndex": index, "endColumnIndex": index + 1}


def push(sheet, df):
    rows = len(df) + 200
    cols = len(LABELS)
    try:
        ws = sheet.worksheet(WORKSHEET)
    except gspread.WorksheetNotFound:
        log.info(f"Created worksheet '{WORKSHEET}'")
        ws = sheets_call(sheet.add_worksheet, title=WORKSHEET, rows=rows, cols=cols)

    if ws.row_count < rows or ws.col_count < cols:
        sheets_call(ws.resize, rows=max(ws.row_count, rows), cols=max(ws.col_count, cols))
    sheets_call(ws.clear)

    sid = ws.id
    # TEXT format on the identifier columns BEFORE the write: set_with_dataframe
    # writes USER_ENTERED, which would strip the leading zeros off LC numbers.
    text_cols = [i for i, label in enumerate(LABELS) if KIND[label] == "id"]
    if text_cols:
        sheets_call(sheet.batch_update, {"requests": [
            {"repeatCell": {"range": col_range(sid, i, len(df) + 1),
                            "cell": {"userEnteredFormat": {"numberFormat": {"type": "TEXT"}}},
                            "fields": "userEnteredFormat.numberFormat"}}
            for i in text_cols]})

    sheets_call(set_with_dataframe, ws, df, row=1, resize=False)

    reqs = [
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": cols},
            "cell": {"userEnteredFormat": {
                "backgroundColor": BLUE,
                "textFormat": {"bold": True, "foregroundColor": {
                    "red": 1, "green": 1, "blue": 1}}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid,
                           "gridProperties": {"frozenRowCount": 1}},
            "fields": "gridProperties.frozenRowCount"}},
        {"setBasicFilter": {"filter": {"range": {
            "sheetId": sid, "startRowIndex": 0, "endRowIndex": len(df) + 1,
            "startColumnIndex": 0, "endColumnIndex": cols}}}},
    ]
    for i, label in enumerate(LABELS):
        pattern = NUMBER_PATTERNS.get(KIND[label])
        if pattern:
            reqs.append({"repeatCell": {
                "range": col_range(sid, i, len(df) + 1),
                "cell": {"userEnteredFormat": {"numberFormat": pattern}},
                "fields": "userEnteredFormat.numberFormat"}})
    sheets_call(sheet.batch_update, {"requests": reqs})
    log.info(f"Wrote {len(df)} rows to '{WORKSHEET}'")


# ========= MAIN ==========
def main():
    as_of = datetime.now(DHAKA).date()
    login()
    df = build(as_of)
    log.info(f"Total {len(df)} rows")

    if os.getenv("RAW_DRY_RUN"):
        out = os.getenv("RAW_DRY_RUN_PATH", "pi_dispatch_raw_dryrun.csv")
        df.to_csv(out, index=False, encoding="utf-8-sig")
        log.info(f"DRY RUN - wrote {out}, no Sheets write")
        return

    client = get_gspread_client()
    sheet = sheets_call(client.open_by_key, SHEET_KEY)
    push(sheet, df)
    log.info("Done")


if __name__ == "__main__":
    main()
