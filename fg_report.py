"""
FG Stock report -> Google Sheets.

Source: operation.details.retrive_data_from_operation_details([[]]), the call the
Odoo FG Store dashboard makes (captured in FG report.har, action 2591).

Six tabs, laid out to match "Fg_Stock AS ON 15-08-2026 (1).xlsx":

  FG Stock                     the dashboard's Excel-export layout (Fg_Stock.xlsx):
                               33 columns, Total row, sorted ascending by oa_id,
                               plus the five derived columns below.
  Full Goods Ready LC RCV      \\
  Partially ready LC RCV        |  the 2x2 split on LC Status x FG Status, each
  Full Goods Ready LC Pending   |  with a banner and a subtotal row, 19 columns.
  Partially ready LC Pending   /
  SUMMERY                      Stock Value by LC/FG status against the DELIVERY
                               DATE ageing buckets, with links to the four tabs.

DERIVED COLUMNS. Reverse-engineered from the user's workbooks and verified to
reproduce every row of both with zero mismatches:

  LC Status  "LC Received" when Invoice No is present, else "LC Pending".
  FG Status  LC status crossed with whether goods are still outstanding, where
             outstanding means the 2dp-rounded Pending Value > 0:
                 LC Received + none outstanding -> Full
                 LC Received + outstanding      -> Partial
                 LC Pending  + none outstanding -> Full OA
                 LC Pending  + outstanding      -> Goods in Production
             Pending Value rather than Pending QTY: the source treats rows with
             Pending QTY 0 but a residual 0.05 as outstanding, and negative
             rounding noise (-0.04) as settled.
  Age        days since goods-in (the endpoint's days_passed), bucketed by Ageing.
  Age DD     days since the ED Date - negative when delivery is not yet due.
  Ageing DD  Age DD bucketed, with a "Delivery not due" band for negative values.
             SUMMERY pivots on this, not on Ageing.

Both ageing scales use the same bands: <=5 / <=10 / <=20 / <=30 / <=60 / <=90 / >90.

Four defects in the source Excel export are deliberately NOT reproduced - see
EXPORT_FIXES.

Target sheet: https://docs.google.com/spreadsheets/d/1YRLVLKbrwXIziBAmtAzgEfEPH64Ycwu57ZZJrDO6aks
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

SHEET_KEY = "1YRLVLKbrwXIziBAmtAzgEfEPH64Ycwu57ZZJrDO6aks"
MASTER_WORKSHEET = "FG Stock"
SUMMARY_WORKSHEET = "SUMMERY"

# Tabs from earlier revisions of this script, superseded by the ones above.
OBSOLETE_WORKSHEETS = [
    "FG Store", "FG Store Details",
    "Goods Ready LC RCV", "Good Pending LC RCV",
    "Goods Ready LC Pending", "Goods Pending LC Pending",
]

ALLOWED_COMPANY_IDS = [2, 1, 3]   # as the dashboard sends: Head Office, Zipper, Metal Trims
DHAKA = timezone(timedelta(hours=6))

EXPORT_FIXES = """\
  1. Total row alignment - the source export writes the totals two columns to the
     left of their headers (the Order QTY total lands under 'Item'). Here each
     total sits under its own column. 'Age' is left blank, as in the source.
  2. LC Number leading zeros - the export coerced them to numbers and lost the
     zeros on 46 of 443 rows ('0000019525120040' -> 19525120040). Written as text.
  3. Date cells - the export stamped every date with a spurious 06:00:20 time
     (a UTC+6 conversion artifact). Written as plain dates, so Age DD comes out
     as whole days instead of the workbook's -50.25023148 style fractions.
  4. Half-up rounding - matches the export's JS toFixed(2) on the cells where
     Python's default banker's rounding would disagree (280.125 -> 280.13)."""

# label, API field, kind. The first 33 are verbatim from Fg_Stock.xlsx; the rest
# are derived (see module docstring).
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
    ("Age DD",          None,                "age"),
    ("Ageing DD",       None,                "text"),
    ("Ageing",          None,                "text"),
    ("LC Status",       None,                "text"),
    ("FG Status",       None,                "text"),
]
LABELS = [label for label, _, _ in COLUMNS]
KIND = {label: kind for label, _, kind in COLUMNS}

# The category tabs: the master's columns minus the ageing/derived extras, in the
# workbook's order.
CATEGORY_LABELS = [
    "OA", "Order Date", "ED Date", "Closing Date", "PI", "Customer", "Buyer",
    "Invoice No", "Invoice Date", "Sales Person", "Sales Team", "Item",
    "Pending QTY", "Pending Value", "Stock QTY", "Stock Value", "Age",
    "LC Status", "FG Status",
]
# Columns carried in each category tab's subtotal row, as in the workbook.
SUBTOTAL_LABELS = ["Pending QTY", "Pending Value", "Stock QTY", "Stock Value"]

# worksheet title, LC Status, FG Status, banner, "Kindly attention". Text verbatim
# from the workbook.
CATEGORIES = [
    ("Full Goods Ready LC RCV", "LC Received", "Full",
     "Max Delivery Date Passed – Full -goods Ready & LC Received, Delivery Pending",
     " @ sales"),
    ("Partially ready LC RCV", "LC Received", "Partial",
     "Partial Goods Ready, Balance Under Production – LC Received",
     " @ Sale & Production"),
    ("Full Goods Ready LC Pending", "LC Pending", "Full OA",
     "Max Delivery Date Passed – Full goods Ready but LC Pending",
     " @ sales"),
    ("Partially ready LC Pending", "LC Pending", "Goods in Production",
     "Partial Goods Ready, Balance Under Production – LC Pending",
     " @ Sale"),
]

AGEING_BUCKETS = [
    (5,            "01-05 DAYS"),
    (10,           "06-10 DAYS"),
    (20,           "11-20 DAYS"),
    (30,           "21-30 DAYS"),
    (60,           "31-60 DAYS"),
    (90,           "61-90 DAYS"),
    (float("inf"), ">90 DAYS"),
]
AGEING_ORDER = [label for _, label in AGEING_BUCKETS]
NOT_DUE = "Delivery not due"
# SUMMERY pivots on Ageing DD, so the not-due band is a column of its own.
SUMMARY_COLUMNS = AGEING_ORDER + [NOT_DUE]
# Parent LC Status followed by its FG Statuses, in the workbook's order.
SUMMARY_GROUPS = [
    ("LC Received", "LC received", ["Full", "Partial"]),
    ("LC Pending",  "LC Pending",  ["Full OA", "Goods in Production"]),
]

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


def ageing_bucket(age):
    if age == "" or age is None:
        return ""
    for limit, label in AGEING_BUCKETS:
        if age <= limit:
            return label
    return AGEING_ORDER[-1]


def ageing_dd_bucket(age_dd):
    if age_dd == "" or age_dd is None:
        return ""
    return NOT_DUE if age_dd < 0 else ageing_bucket(age_dd)


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
    """Six tabs means a burst of Sheets writes; back off on rate limits."""
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


# ========= TABLES ==========
def build_master(records, as_of):
    """Header + Total + data, the Fg_Stock.xlsx shape plus the derived columns."""
    data = []
    for rec in records:
        row = {}
        for label, key, kind in COLUMNS:
            if key is None:
                continue
            v = rec.get(key)
            if kind == "text":
                row[label] = clean_text(v)
            elif kind == "date":
                row[label] = clean_date(v)
            elif kind == "age":
                row[label] = int(v) if isinstance(v, (int, float)) else ""
            else:
                row[label] = round2(v)

        # Days since the expected delivery date; negative until it falls due.
        try:
            row["Age DD"] = (as_of - date.fromisoformat(row["ED Date"])).days
        except ValueError:
            row["Age DD"] = ""
        row["Ageing DD"] = ageing_dd_bucket(row["Age DD"])
        row["Ageing"] = ageing_bucket(row["Age"])

        # Evaluated on the rounded Pending Value so that residual fractions and
        # negative noise classify the way the source workbook does.
        row["LC Status"] = "LC Received" if row["Invoice No"] else "LC Pending"
        outstanding = (row["Pending Value"] or 0) > 0
        if row["LC Status"] == "LC Received":
            row["FG Status"] = "Partial" if outstanding else "Full"
        else:
            row["FG Status"] = "Goods in Production" if outstanding else "Full OA"
        data.append(row)

    total = {label: "" for label in LABELS}
    total["OA"] = "Total"
    for label, key, kind in COLUMNS:
        if kind in ("qty", "value"):
            total[label] = round2(sum(rec.get(key) or 0 for rec in records))

    return pd.DataFrame([total] + data, columns=LABELS)


def build_category(master_df, lc_status, fg_status):
    """One of the four splits: 19 columns, oldest stock first."""
    body = master_df.iloc[1:]
    sel = body[(body["LC Status"] == lc_status) & (body["FG Status"] == fg_status)]
    sel = sel.sort_values("Age", ascending=False, kind="stable")
    return sel[CATEGORY_LABELS].reset_index(drop=True)


def build_summary(master_df, tab_gids):
    """Stock Value by LC Status / FG Status against the DELIVERY DATE ageing
    buckets, mirroring the workbook's SUMMERY pivot."""
    body = master_df.iloc[1:]
    cat_by_fg = {fg: (title, banner, attn) for title, _, fg, banner, attn in CATEGORIES}

    def line(sel, label, particulars, attn="", tab=None):
        cells = [round2(sel[sel["Ageing DD"] == c]["Stock Value"].sum())
                 for c in SUMMARY_COLUMNS]
        link = ""
        if tab and tab in tab_gids:
            link = f'=HYPERLINK("#gid={tab_gids[tab]}","{tab}")'
        return ([particulars, label] + cells
                + [round2(sel["Stock Value"].sum()), attn, link])

    rows = []
    for lc, particulars, fg_list in SUMMARY_GROUPS:
        rows.append(line(body[body["LC Status"] == lc], lc, particulars))
        for fg in fg_list:
            title, banner, attn = cat_by_fg[fg]
            rows.append(line(body[body["FG Status"] == fg], fg, banner, attn, title))
    rows.append(line(body, "Grand Total", "Grand Total"))

    header = (["Particulars", "Row Labels"] + SUMMARY_COLUMNS
              + ["Grand Total", "Kindly attention", "sheet link"])
    return pd.DataFrame(rows, columns=header)


SUMMARY_KINDS = {"Particulars": "text", "Row Labels": "text",
                 "Kindly attention": "text", "sheet link": "formula"}


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


NUMBER_PATTERNS = {"date": {"type": "DATE", "pattern": "yyyy-mm-dd"},
                   "qty":  {"type": "NUMBER", "pattern": "#,##0.##"},
                   "value": {"type": "NUMBER", "pattern": "#,##0.00"},
                   "age":  {"type": "NUMBER", "pattern": "0"}}

BLUE = {"red": 0.10, "green": 0.31, "blue": 0.58}
PALE = {"red": 0.85, "green": 0.89, "blue": 0.96}


def get_worksheet(sheet, title, rows, cols):
    try:
        return sheet.worksheet(title)
    except gspread.WorksheetNotFound:
        log.info(f"Created worksheet '{title}'")
        return sheets_call(sheet.add_worksheet, title=title, rows=rows, cols=cols)


def push(sheet, title, df, kinds, header_row=1, pre_cells=None,
         bold_rows=(), freeze=None):
    """Write one tab.

    header_row  1-indexed row the column headers land on; data follows.
    pre_cells   [(a1_range, values_2d)] written above the frame.
    bold_rows   1-indexed rows to bold and shade (e.g. a Total or subtotal row).
    """
    ws = get_worksheet(sheet, title, max(1000, len(df) + header_row + 50),
                       max(40, len(df.columns) + 2))

    need_rows = max(len(df) + header_row + 50, 200)
    need_cols = max(len(df.columns) + 2, 40)
    if ws.row_count < need_rows or ws.col_count < need_cols:
        sheets_call(ws.resize, rows=max(ws.row_count, need_rows),
                    cols=max(ws.col_count, need_cols))

    sheets_call(ws.clear)
    sid = ws.id

    def fmt(idx, number_format):
        return {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1,
                      "startColumnIndex": idx, "endColumnIndex": idx + 1},
            "cell": {"userEnteredFormat": {"numberFormat": number_format}},
            "fields": "userEnteredFormat.numberFormat"}}

    # clear() wipes values but not formatting; reset so last run's formats can't
    # reinterpret this run's values.
    sheets_call(sheet.batch_update, {"requests": [
        {"updateCells": {"range": {"sheetId": sid}, "fields": "userEnteredFormat"}}]})

    # Text format must land BEFORE the values: USER_ENTERED honours an existing
    # TEXT format, and that is what keeps the leading zeros on LC numbers. The
    # 'formula' kind is skipped so HYPERLINK() is not stored as literal text.
    text_reqs = [fmt(i, {"type": "TEXT"})
                 for i, c in enumerate(df.columns) if kinds.get(c) == "text"]
    if text_reqs:
        sheets_call(sheet.batch_update, {"requests": text_reqs})

    for a1, values in (pre_cells or []):
        sheets_call(ws.update, values, a1, value_input_option="USER_ENTERED")

    sheets_call(set_with_dataframe, ws, df, row=header_row, resize=False)

    reqs = [fmt(i, NUMBER_PATTERNS[kinds[c]])
            for i, c in enumerate(df.columns) if kinds.get(c) in NUMBER_PATTERNS]
    reqs.append({"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": header_row - 1, "endRowIndex": header_row},
        "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": BLUE, "horizontalAlignment": "CENTER"}},
        "fields": "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment)"}})
    for r in bold_rows:
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": r - 1, "endRowIndex": r},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "backgroundColor": PALE}},
            "fields": "userEnteredFormat(textFormat,backgroundColor)"}})
    reqs.append({"updateSheetProperties": {
        "properties": {"sheetId": sid,
                       "gridProperties": {"frozenRowCount": freeze if freeze is not None else header_row}},
        "fields": "gridProperties.frozenRowCount"}})
    sheets_call(sheet.batch_update, {"requests": reqs})

    log.info(f"  '{title}': {len(df)} rows x {len(df.columns)} cols")
    return ws


def push_category(sheet, title, df, banner):
    """Banner on row 1, subtotals on row 2, headers on row 3, data from row 4."""
    sub = [""] * len(CATEGORY_LABELS)
    for label in SUBTOTAL_LABELS:
        sub[CATEGORY_LABELS.index(label)] = round2(
            pd.to_numeric(df[label], errors="coerce").fillna(0).sum())
    return push(sheet, title, df, KIND, header_row=3,
                pre_cells=[("A1", [[banner]]), ("A2", [sub])],
                bold_rows=(1, 2), freeze=3)


def push_summary(sheet, df, as_of):
    title = (f" Finish Goods status as on {as_of.day} {as_of:%B %Y} "
             f"( Ageing delivery date wise)")
    kinds = {c: SUMMARY_KINDS.get(c, "value") for c in df.columns}
    return push(sheet, SUMMARY_WORKSHEET, df, kinds, header_row=5,
                pre_cells=[("A2", [[title]]), ("B4", [["Sum of Stock Value", "Column Labels"]])],
                bold_rows=(2, 6, 9, 12), freeze=5)


def drop_obsolete(sheet):
    existing = {ws.title: ws for ws in sheet.worksheets()}
    for title in OBSOLETE_WORKSHEETS:
        ws = existing.get(title)
        if ws is not None and len(existing) > 1:
            sheets_call(sheet.del_worksheet, ws)
            existing.pop(title)
            log.info(f"Removed obsolete worksheet '{title}'")


# ========= MAIN ==========
def main():
    missing = [n for n, v in [("ODOO_URL", ODOO_URL), ("ODOO_DB", DB),
                              ("ODOO_USERNAME", USERNAME), ("ODOO_PASSWORD", PASSWORD)] if not v]
    if missing:
        raise SystemExit(f"Missing env vars: {', '.join(missing)}")

    login()
    as_of = datetime.now(DHAKA).date()
    master = build_master(fetch_stock(), as_of)
    body = master.iloc[1:]

    categories = [(title, banner, build_category(master, lc, fg))
                  for title, lc, fg, banner, _ in CATEGORIES]

    totals = master.iloc[0]
    log.info("Totals: " + "  ".join(
        f"{lbl}={totals[lbl]:,.2f}" for lbl, _, kind in COLUMNS if kind in ("qty", "value")))
    log.info("Split: " + "  ".join(f"{t}={len(d)}" for t, _, d in categories)
             + f"  (sum={sum(len(d) for _, _, d in categories)} of {len(body)})")
    log.info("Ageing DD: " + "  ".join(
        f"{c}={(body['Ageing DD'] == c).sum()}" for c in SUMMARY_COLUMNS))

    covered = sum(len(d) for _, _, d in categories)
    if covered != len(body):
        log.info(f"WARNING: {len(body) - covered} rows fall outside the four categories")
    no_ed = (body["Age DD"] == "").sum()
    if no_ed:
        log.info(f"NOTE: {no_ed} rows have no ED Date, so Age DD / Ageing DD are blank")

    if os.getenv("FG_REPORT_DRY_RUN"):
        out = os.getenv("FG_REPORT_DRY_RUN_PATH", "fg_stock_dryrun.csv")
        master.to_csv(out, index=False)
        for title, _, d in categories:
            d.to_csv(out.replace(".csv", f" - {title}.csv"), index=False)
        build_summary(master, {}).to_csv(out.replace(".csv", " - SUMMERY.csv"), index=False)
        log.info(f"DRY RUN: wrote {out} and 5 companions, skipped Google Sheets")
        return

    client = get_gspread_client()
    sheet = sheets_call(client.open_by_key, SHEET_KEY)

    push(sheet, MASTER_WORKSHEET, master, KIND, header_row=1, bold_rows=(2,), freeze=2)
    # Category tabs first: SUMMERY links to them by sheet id.
    tab_gids = {}
    for title, banner, d in categories:
        tab_gids[title] = push_category(sheet, title, d, banner).id
    push_summary(sheet, build_summary(master, tab_gids), as_of)
    # Drop last: the tabs just written guarantee the spreadsheet still has a
    # sheet left, so no delete can hit the last-tab guard.
    drop_obsolete(sheet)

    log.info(f"Done at {datetime.now(DHAKA):%Y-%m-%d %H:%M} Asia/Dhaka")


if __name__ == "__main__":
    main()
