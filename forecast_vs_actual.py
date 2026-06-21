import requests
import json
import base64
import logging
import sys
import os

if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding and sys.stderr.encoding.lower() != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')

import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2 import service_account
import pandas as pd
import time
from dotenv import load_dotenv
from requests.exceptions import RequestException

load_dotenv()
logging.basicConfig(stream=sys.stdout, level=logging.INFO)
log = logging.getLogger()

# ========= CONFIG ==========
ODOO_URL = os.getenv("ODOO_URL")
DB = os.getenv("ODOO_DB")
USERNAME = os.getenv("ODOO_USERNAME")
PASSWORD = os.getenv("ODOO_PASSWORD")

SHEET_KEY = "1coN06mZ9uLBn1JnSyNqLwYhpl2-fJsuT85xwuNI0iy8"
FORECAST_WORKSHEET_NAME = "Forecast"

# Dump the RAW rolling.forecast lines (the data BEFORE the dashboard's server-side
# GROUP BY) so each dimension is its own column and you can pivot/group yourself.
# Pull for both companies; the Company column distinguishes them.
COMPANY_IDS = [3, 1]        # 3 = "Metal Trims", 1 = "Zipper"

session = requests.Session()
USER_ID = None

# ========= GOOGLE SHEETS CLIENT ==========
def get_gspread_client():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for sa_path in [
        os.path.join(script_dir, "service_account.json"),
        "service_account.json",
    ]:
        if os.path.exists(sa_path):
            return gspread.service_account(filename=sa_path)

    # GitHub Actions style: full JSON secret stored in GOOGLE_SHEET_CRED_JSON
    creds_json = os.getenv("GOOGLE_SHEET_CRED_JSON")
    if creds_json:
        creds_dict = json.loads(creds_json)
        scopes = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)

    creds_raw = os.getenv("GOOGLE_CREDS_BASE64")
    if creds_raw:
        creds_dict = None
        try:
            creds_dict = json.loads(creds_raw.strip())
        except json.JSONDecodeError:
            pass
        if creds_dict is None:
            try:
                padded = creds_raw.strip() + '=' * (-len(creds_raw.strip()) % 4)
                creds_dict = json.loads(base64.b64decode(padded).decode("utf-8"))
            except Exception:
                pass
        if creds_dict is None:
            raise Exception("GOOGLE_CREDS_BASE64 is neither valid JSON nor valid base64-encoded JSON")
        scopes = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)

    raise Exception("No Google credentials found.")

# ========= RETRY LOGIC ==========
def retry_request(method, url, max_retries=3, backoff=3, **kwargs):
    for attempt in range(1, max_retries + 1):
        try:
            r = method(url, **kwargs)
            r.raise_for_status()
            return r
        except RequestException as e:
            print(f"⚠️ Attempt {attempt} failed: {e}")
            if attempt < max_retries:
                print(f"⏳ Retrying in {backoff} seconds...")
                time.sleep(backoff)
            else:
                print("❌ All retry attempts failed.")
                raise

# ========= LOGIN ==========
def login():
    global USER_ID
    payload = {"jsonrpc": "2.0", "params": {"db": DB, "login": USERNAME, "password": PASSWORD}}
    r = retry_request(session.post, f"{ODOO_URL}/web/session/authenticate", json=payload)
    result = r.json().get("result")
    if result and "uid" in result:
        USER_ID = result["uid"]
        print(f"✅ Logged in (uid={USER_ID})")
        return result
    else:
        raise Exception("❌ Login failed")

# ========= SWITCH COMPANY ==========
def switch_company(company_id):
    if USER_ID is None:
        raise Exception("User not logged in yet")
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "res.users",
            "method": "write",
            "args": [[USER_ID], {"company_id": company_id}],
            "kwargs": {"context": {"allowed_company_ids": [company_id], "company_id": company_id}},
        },
    }
    r = retry_request(session.post, f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    if "error" in r.json():
        print(f"❌ Failed to switch to company {company_id}: {r.json()['error']}")
        return False
    print(f"🔄 Session switched to company {company_id}")
    return True

# ========= COMPANY NAME ==========
def fetch_company_name(company_id):
    ctx = {"lang": "en_US", "tz": "Asia/Dhaka", "uid": USER_ID,
           "allowed_company_ids": [company_id]}
    payload = {
        "jsonrpc": "2.0", "method": "call",
        "params": {
            "model": "res.company", "method": "search_read", "args": [],
            "kwargs": {"domain": [["id", "=", company_id]], "fields": ["id", "name"], "context": ctx},
        },
    }
    r = retry_request(session.post, f"{ODOO_URL}/web/dataset/call_kw/res.company/search_read", json=payload)
    res = r.json().get("result") or []
    return res[0]["name"] if res else str(company_id)

# ========= FETCH FORECAST HEADERS (month + line ids) ==========
def fetch_forecast_headers(company_id):
    """
    Read rolling.forecast headers for the company. Each header is one
    salesperson/team projection for a month; next_month_line_ids are the raw
    detail lines. Returns list of header dicts.
    """
    ctx = {"lang": "en_US", "tz": "Asia/Dhaka", "uid": USER_ID,
           "allowed_company_ids": [company_id]}
    out = []
    offset = 0
    page = 2000
    while True:
        payload = {
            "jsonrpc": "2.0", "method": "call",
            "params": {
                "model": "rolling.forecast", "method": "search_read", "args": [],
                "kwargs": {
                    "domain": [["company_id", "=", company_id], ["next_month", "!=", False]],
                    "fields": ["id", "next_month", "next_month_label", "next_month_line_ids"],
                    "offset": offset, "limit": page, "order": "next_month DESC",
                    "context": ctx,
                },
            },
        }
        r = retry_request(session.post,
                          f"{ODOO_URL}/web/dataset/call_kw/rolling.forecast/search_read",
                          json=payload)
        recs = r.json().get("result") or []
        out.extend(recs)
        if len(recs) < page:
            break
        offset += page
    return out

# ========= FETCH RAW FORECAST LINES ==========
# Dimension/measure fields pulled from each rolling.forecast.line so you can
# group by any of them yourself in the sheet.
LINE_FIELDS = [
    "id", "qty", "avg_price", "total_price", "achived", "achieved_value",
    "item_achived", "account_type", "type", "classification", "segments",
    "sales_person_region", "forecast_product_id", "item_category",
    "customer_name", "customer_group", "buyer", "brand_group",
    "salesperson_id", "sales_team_id", "division_state", "currency_id",
]

def fetch_forecast_lines(company_id, line_ids):
    """Batch-read rolling.forecast.line records by id with all dimension fields."""
    ctx = {"lang": "en_US", "tz": "Asia/Dhaka", "uid": USER_ID,
           "allowed_company_ids": [company_id]}
    out = []
    CHUNK = 2000
    for i in range(0, len(line_ids), CHUNK):
        chunk = line_ids[i:i + CHUNK]
        payload = {
            "jsonrpc": "2.0", "method": "call",
            "params": {
                "model": "rolling.forecast.line", "method": "read",
                "args": [chunk, LINE_FIELDS],
                "kwargs": {"context": ctx},
            },
        }
        r = retry_request(session.post,
                          f"{ODOO_URL}/web/dataset/call_kw/rolling.forecast.line/read",
                          json=payload)
        body = r.json()
        if "error" in body:
            raise Exception(f"line read failed: {body['error'].get('message') or body['error']}")
        out.extend(body.get("result") or [])
    return out

def _m2o_name(v):
    """Odoo many2one comes back as [id, name] or False -> return name or ''."""
    return v[1] if isinstance(v, (list, tuple)) and len(v) == 2 else ""

# ========= BUILD RAW FORECAST ROWS (one row per forecast line) ==========
def build_forecast_rows(company_id, company_name, headers):
    """
    Flatten raw rolling.forecast.line records into one row each, tagged with
    Company + Month. Every dimension is its own column so you can group by
    Product / Customer / Brand / Salesperson / Team / Division yourself.

    Months emitted newest → oldest (headers are pre-sorted that way).
    """
    # month per line id, from the headers
    line_month = {}   # line_id -> (month, month_label)
    all_line_ids = []
    for h in headers:
        m = h.get("next_month")
        ml = h.get("next_month_label") or m
        for lid in h.get("next_month_line_ids") or []:
            if lid not in line_month:
                line_month[lid] = (m, ml)
                all_line_ids.append(lid)

    print(f"   {company_name}: {len(headers)} headers, {len(all_line_ids)} forecast lines")
    lines = fetch_forecast_lines(company_id, all_line_ids)

    out = []
    for ln in lines:
        month, month_label = line_month.get(ln["id"], ("", ""))
        out.append({
            "Company":        company_name,
            "Month":          month,
            "Month Label":    month_label,
            "Product":        _m2o_name(ln.get("item_category")),
            "Forecast Product": _m2o_name(ln.get("forecast_product_id")),
            "Customer":       _m2o_name(ln.get("customer_name")),
            "Customer Group": _m2o_name(ln.get("customer_group")),
            "Brand":          _m2o_name(ln.get("buyer")),
            "Brand Group":    _m2o_name(ln.get("brand_group")),
            "Salesperson":    _m2o_name(ln.get("salesperson_id")),
            "Sales Team":     _m2o_name(ln.get("sales_team_id")),
            "Division":       ln.get("division_state") or "",
            "Region":         ln.get("sales_person_region") or "",
            "Account Type":   ln.get("account_type") or "",
            "Type":           ln.get("type") or "",
            "Classification": ln.get("classification") or "",
            "Segments":       ln.get("segments") or "",
            "Currency":       _m2o_name(ln.get("currency_id")),
            "Forecast QTY":   ln.get("qty", 0.0),
            "Avg Price":      ln.get("avg_price", 0.0),
            "Forecast Value": ln.get("total_price", 0.0),
            "Achieved QTY":   ln.get("achived", 0.0),
            "Achieved Value": ln.get("achieved_value", 0.0),
        })
    # newest month first (headers already DESC, but be explicit / stable)
    out.sort(key=lambda r: r["Month"], reverse=True)
    return out

# Output columns, in order. Dimensions first (group by any of these), then
# measures. Keep this in sync with build_forecast_rows().
COLS = ["Company", "Month", "Month Label", "Product", "Forecast Product",
        "Customer", "Customer Group", "Brand", "Brand Group", "Salesperson",
        "Sales Team", "Division", "Region", "Account Type", "Type",
        "Classification", "Segments", "Currency",
        "Forecast QTY", "Avg Price", "Forecast Value", "Achieved QTY", "Achieved Value"]

# ========= MAIN ==========
if __name__ == "__main__":
    login()

    forecast_rows = []
    for cid in COMPANY_IDS:
        if not switch_company(cid):
            print(f"⚠️ Skipping company {cid} (could not switch)")
            continue
        company_name = fetch_company_name(cid)
        headers = fetch_forecast_headers(cid)
        if not headers:
            print(f"⚠️ {company_name}: no forecast headers, skipping")
            continue
        months = sorted({h.get("next_month") for h in headers if h.get("next_month")})
        print(f"🏢 {company_name} (id={cid}): {len(headers)} headers, "
              f"months {months[0]} → {months[-1]}")
        forecast_rows.extend(build_forecast_rows(cid, company_name, headers))

    if not forecast_rows:
        print("❌ No forecast rows fetched for any company.")
        sys.exit(1)
    print(f"📊 Total forecast lines: {len(forecast_rows)}")

    df = pd.DataFrame(forecast_rows, columns=COLS)

    # ========= PUSH TO GOOGLE SHEETS =========
    try:
        client = get_gspread_client()
        sheet = client.open_by_key(SHEET_KEY)
        try:
            ws = sheet.worksheet(FORECAST_WORKSHEET_NAME)
        except gspread.WorksheetNotFound:
            ws = sheet.add_worksheet(title=FORECAST_WORKSHEET_NAME,
                                     rows=max(1000, len(df) + 50),
                                     cols=max(15, len(df.columns) + 2))
        ws.clear()
        # clear() wipes values but NOT cell formatting; reset any stale formatting so
        # numeric columns don't render oddly (e.g. numbers shown as dates).
        try:
            ws.spreadsheet.batch_update({
                "requests": [{
                    "updateCells": {
                        "range": {"sheetId": ws.id},
                        "fields": "userEnteredFormat",
                    }
                }]
            })
        except Exception as fmt_err:
            print(f"⚠️ Could not clear Forecast formatting: {fmt_err}")

        set_with_dataframe(ws, df)

        # Pin the measure columns to a plain number format. They are the last 5
        # columns of COLS: Forecast QTY, Avg Price, Forecast Value, Achieved QTY,
        # Achieved Value -> 0-based indices 18..22.
        try:
            number_fmt = {"numberFormat": {"type": "NUMBER", "pattern": "0.##########"}}
            requests = []
            for col_idx in range(len(COLS) - 5, len(COLS)):
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": 1,  # skip header
                            "startColumnIndex": col_idx,
                            "endColumnIndex": col_idx + 1,
                        },
                        "cell": {"userEnteredFormat": number_fmt},
                        "fields": "userEnteredFormat.numberFormat",
                    }
                })
            ws.spreadsheet.batch_update({"requests": requests})
        except Exception as fmt_err:
            print(f"⚠️ Could not set Forecast number formats: {fmt_err}")

        print(f"✅ Forecast pasted to Google Sheets → '{FORECAST_WORKSHEET_NAME}' "
              f"({len(df)} raw forecast lines)")
    except Exception as e:
        import traceback
        print(f"❌ Error while pasting to Google Sheets: {e}")
        traceback.print_exc()
