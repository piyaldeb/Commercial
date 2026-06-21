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

# Filter values observed in the "Forecast VS Actual" dashboard HAR:
#   retrieve_unified_performance_dashboard(company_id, month, group_by, only_forecast, type)
#   = (3, "2026-08", "customer", False, "local_foreign")
# Pull the forecast for both companies; the Company column distinguishes them.
COMPANY_IDS = [3, 1]        # 3 = "Metal Trims", 1 = "Zipper"
ONLY_FORECAST = False       # "Only Forecast" checkbox unticked
FORECAST_TYPE = "local_foreign"  # "Local & Foreign"

# Every group-by breakdown the dashboard dropdown offers, as (api_key, label).
# Verified live against the endpoint — each key returns a genuinely distinct
# breakdown. NOTE the product breakdown's API key is "item" (SHANK BUTTON,
# ALLOY, RIVET...); the literal key "product" silently falls back to customer.
GROUP_BYS = [
    ("item",        "By Product"),
    ("salesperson", "By Salesperson"),
    ("team",        "By Team"),
    ("division",    "By Division"),
    ("customer",    "By Customer"),
    ("brand",       "By Brand"),
]

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
    r = retry_request(session.post, f"{ODOO_URL}/web/dataset/call_kw/rolling.forecast/search_read", json=payload)
    recs = r.json().get("result") or []
    seen = {}
    for rec in recs:
        m = rec.get("next_month")
        if m and m not in seen:
            seen[m] = rec.get("next_month_label") or m
    # chronological order (next_month is YYYY-MM, sorts correctly as string)
    return [{"next_month": m, "next_month_label": seen[m]} for m in sorted(seen)]

# ========= FETCH DASHBOARD (Forecast vs Actual) ==========
def fetch_forecast_dashboard(company_id, month, group_by, only_forecast, ftype):
    """
    Call rolling.forecast.retrieve_unified_performance_dashboard.
    Returns dict: {rows: [...], totals: {...}, month_label: "..."}.
    Each row: name, forecast_qty, forecast_value, oa_qty, oa_value, growth_pct.
    """
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

# ========= BUILD FORECAST ROWS (long format, all months × all group-bys) ==========
def build_forecast_rows(company_id, company_name, months, group_bys, only_forecast, ftype):
    """
    One output row per (month, group-by, entry). Every dashboard filter is a
    column, and the breakdowns are stacked with a "Group By" column so the sheet
    behaves like the dashboard: filter Group By = "By Salesperson" (or Brand /
    Team / Division / Customer) to see that exact breakdown.

    Columns: Company, Type, Group By, Only Forecast, Month, Month Label, Name,
             Forecast QTY, Forecast Value, OA QTY, OA Value, Growth %

    Months are emitted newest → oldest.
    """
    type_label = {
        "local_foreign": "Local & Foreign",
        "local": "Local",
        "foreign": "Foreign",
    }.get(ftype, ftype)

    # newest month first
    months_sorted = sorted(months, key=lambda m: m["next_month"], reverse=True)

    out = []
    for m in months_sorted:
        month = m["next_month"]
        month_label = m["next_month_label"]
        for group_key, group_label in group_bys:
            data = fetch_forecast_dashboard(company_id, month, group_key, only_forecast, ftype)
            rows = data.get("rows") or []
            label_from_resp = data.get("month_label") or month_label
            for r in rows:
                out.append({
                    "Company":        company_name,
                    "Type":           type_label,
                    "Group By":       group_label,
                    "Only Forecast":  "Yes" if only_forecast else "No",
                    "Month":          month,
                    "Month Label":    label_from_resp,
                    "Name":           r.get("name", ""),
                    "Forecast QTY":   r.get("forecast_qty", 0.0),
                    "Forecast Value": r.get("forecast_value", 0.0),
                    "OA QTY":         r.get("oa_qty", 0.0),
                    "OA Value":       r.get("oa_value", 0.0),
                    "Growth %":       r.get("growth_pct", 0.0),
                })
            print(f"📋 {month_label} / {group_label}: {len(rows)} rows")
    return out

# ========= MAIN ==========
if __name__ == "__main__":
    login()

    print(f"🧩 group-bys: {', '.join(lbl for _, lbl in GROUP_BYS)}")
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
        forecast_rows.extend(build_forecast_rows(
            cid, company_name, months, GROUP_BYS, ONLY_FORECAST, FORECAST_TYPE))

    if not forecast_rows:
        print("❌ No forecast rows fetched for any company.")
        sys.exit(1)
    print(f"📊 Total forecast rows: {len(forecast_rows)}")

    cols = ["Company", "Type", "Group By", "Only Forecast", "Month", "Month Label",
            "Name", "Forecast QTY", "Forecast Value", "OA QTY", "OA Value", "Growth %"]
    df = pd.DataFrame(forecast_rows, columns=cols)

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

        # Pin numeric columns to a plain number format. 0-based indices:
        #   Forecast QTY=7(H), Forecast Value=8(I), OA QTY=9(J), OA Value=10(K), Growth %=11(L)
        try:
            number_fmt = {"numberFormat": {"type": "NUMBER", "pattern": "0.##########"}}
            requests = []
            for col_idx in (7, 8, 9, 10, 11):
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
              f"({len(df)} rows; {len(months)} months)")
    except Exception as e:
        import traceback
        print(f"❌ Error while pasting to Google Sheets: {e}")
        traceback.print_exc()
