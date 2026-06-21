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

# ========= FETCH AVAILABLE FORECAST MONTHS ==========
def fetch_forecast_months(company_id):
    """Return ordered list of {next_month, next_month_label} distinct values."""
    ctx = {"lang": "en_US", "tz": "Asia/Dhaka", "uid": USER_ID,
           "allowed_company_ids": [company_id]}
    payload = {
        "jsonrpc": "2.0", "method": "call",
        "params": {
            "model": "rolling.forecast", "method": "search_read", "args": [],
            "kwargs": {
                "domain": [["company_id", "=", company_id], ["next_month", "!=", False]],
                "fields": ["next_month", "next_month_label"],
                "context": ctx,
            },
        },
    }
    r = retry_request(session.post,
                      f"{ODOO_URL}/web/dataset/call_kw/rolling.forecast/search_read", json=payload)
    recs = r.json().get("result") or []
    seen = {}
    for rec in recs:
        m = rec.get("next_month")
        if m and m not in seen:
            seen[m] = rec.get("next_month_label") or m
    return [{"next_month": m, "next_month_label": seen[m]} for m in sorted(seen)]

# ========= FETCH FORECAST-VS-ACTUAL (dashboard) ==========
# group_by = "customer": one row per customer. The dashboard pairs the forecast
# with the actual OA-released figures (its "OA Released" columns) — that is the
# real forecast-vs-actual comparison. (achieved_value on the raw line is a broken
# per-header rollup, so we use the dashboard instead.)
GROUP_BY = "customer"
ONLY_FORECAST = False
FORECAST_TYPE = "local_foreign"

def fetch_forecast_dashboard(company_id, month, group_by, only_forecast, ftype):
    """Return {rows, totals, month_label} for one month/breakdown."""
    ctx = {"lang": "en_US", "tz": "Asia/Dhaka", "uid": USER_ID,
           "allowed_company_ids": [company_id]}
    payload = {
        "jsonrpc": "2.0", "method": "call",
        "params": {
            "model": "rolling.forecast",
            "method": "retrieve_unified_performance_dashboard",
            "args": [company_id, month, group_by, only_forecast, ftype],
            "kwargs": {"context": ctx},
        },
    }
    r = retry_request(
        session.post,
        f"{ODOO_URL}/web/dataset/call_kw/rolling.forecast/retrieve_unified_performance_dashboard",
        json=payload,
    )
    body = r.json()
    if "error" in body:
        raise Exception(f"dashboard fetch failed for {month}: "
                        f"{body['error'].get('message') or body['error']}")
    return body.get("result") or {"rows": [], "totals": {}, "month_label": month}

# ========= BUILD FORECAST-VS-ACTUAL ROWS ==========
def build_forecast_rows(company_id, company_name, months):
    """
    One row per (month, customer): Forecast QTY/Value vs Actual (OA Released)
    QTY/Value, plus Achievement %. Months emitted newest → oldest.
    """
    months_sorted = sorted(months, key=lambda m: m["next_month"], reverse=True)
    out = []
    for m in months_sorted:
        month = m["next_month"]
        data = fetch_forecast_dashboard(company_id, month, GROUP_BY, ONLY_FORECAST, FORECAST_TYPE)
        rows = data.get("rows") or []
        month_label = data.get("month_label") or m["next_month_label"]
        for r in rows:
            fval = r.get("forecast_value") or 0.0
            oaval = r.get("oa_value") or 0.0
            out.append({
                "Company":         company_name,
                "Month":           month,
                "Month Label":     month_label,
                "Customer":        r.get("name", ""),
                "Forecast QTY":    r.get("forecast_qty") or 0.0,
                "Forecast Value":  fval,
                "Actual QTY":      r.get("oa_qty") or 0.0,
                "Actual Value":    oaval,
                # Achievement = actual / forecast, as a fraction (format as % in sheet)
                "Achievement %":   (oaval / fval) if fval else 0.0,
            })
        print(f"📋 {month_label}: {len(rows)} customers")
    return out

# Output columns, in order. Forecast vs Actual, by customer.
COLS = ["Company", "Month", "Month Label", "Customer",
        "Forecast QTY", "Forecast Value", "Actual QTY", "Actual Value",
        "Achievement %"]

# ========= MAIN ==========
if __name__ == "__main__":
    login()

    forecast_rows = []
    for cid in COMPANY_IDS:
        if not switch_company(cid):
            print(f"⚠️ Skipping company {cid} (could not switch)")
            continue
        company_name = fetch_company_name(cid)
        months = fetch_forecast_months(cid)
        if not months:
            print(f"⚠️ {company_name}: no forecast months, skipping")
            continue
        print(f"🏢 {company_name} (id={cid}): {len(months)} months "
              f"{months[0]['next_month']} → {months[-1]['next_month']}")
        forecast_rows.extend(build_forecast_rows(cid, company_name, months))

    if not forecast_rows:
        print("❌ No forecast rows fetched for any company.")
        sys.exit(1)
    print(f"📊 Total forecast-vs-actual rows: {len(forecast_rows)}")

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

        # Format the measure columns. Forecast QTY/Value + Actual QTY/Value are
        # plain numbers (COLS idx 4..7); Achievement % is a percent (idx 8).
        try:
            number_fmt = {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.##"}}
            pct_fmt = {"numberFormat": {"type": "PERCENT", "pattern": "0.0%"}}

            def _fmt_req(idx, fmt):
                return {"repeatCell": {
                    "range": {"sheetId": ws.id, "startRowIndex": 1,
                              "startColumnIndex": idx, "endColumnIndex": idx + 1},
                    "cell": {"userEnteredFormat": fmt},
                    "fields": "userEnteredFormat.numberFormat"}}

            requests = [_fmt_req(i, number_fmt) for i in range(4, 8)]  # qty/value cols
            requests.append(_fmt_req(8, pct_fmt))                       # Achievement %
            ws.spreadsheet.batch_update({"requests": requests})
        except Exception as fmt_err:
            print(f"⚠️ Could not set Forecast number formats: {fmt_err}")

        print(f"✅ Forecast pasted to Google Sheets → '{FORECAST_WORKSHEET_NAME}' "
              f"({len(df)} forecast-vs-actual rows)")
    except Exception as e:
        import traceback
        print(f"❌ Error while pasting to Google Sheets: {e}")
        traceback.print_exc()
