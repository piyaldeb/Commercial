"""
Pull every OA listed under the Payment Terms section of the FG Stock Dashboard
and write a flat OA list to the 'FG store-Agieng Wise' worksheet on Google Sheets.

Calls fg.stock.dashboard.get_drilldown_data once per payment-term bucket
(same call the Odoo web UI makes when you click a value).
"""

import os
import sys
import json
import base64
import logging

import requests
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2 import service_account
from dotenv import load_dotenv

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")

load_dotenv()
logging.basicConfig(stream=sys.stdout, level=logging.INFO, format="%(message)s")
log = logging.getLogger()

ODOO_URL = os.getenv("ODOO_URL", "").rstrip("/")
DB = os.getenv("ODOO_DB")
USERNAME = os.getenv("ODOO_USERNAME")
PASSWORD = os.getenv("ODOO_PASSWORD")

SHEET_KEY = "1coN06mZ9uLBn1JnSyNqLwYhpl2-fJsuT85xwuNI0iy8"
WORKSHEET_NAME = "FG store-Agieng Wise"

# Same filter the browser sent in the HAR.
DASHBOARD_FILTER = {
    "company_id": "all",
    "salesperson_id": None,
    "team_id": None,
    "company_ids": [1, 3],
}
ALLOWED_COMPANY_IDS = [1, 4, 3]
METRIC = "Total_Value"

session = requests.Session()


def jsonrpc(model: str, method: str, args, kwargs=None) -> dict:
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": model,
            "method": method,
            "args": args,
            "kwargs": kwargs or {},
        },
    }
    r = session.post(
        f"{ODOO_URL}/web/dataset/call_kw/{model}/{method}",
        json=payload,
        timeout=120,
    )
    r.raise_for_status()
    data = r.json()
    if "error" in data:
        raise RuntimeError(json.dumps(data["error"], indent=2))
    return data["result"]


def authenticate():
    r = session.post(
        f"{ODOO_URL}/web/session/authenticate",
        json={
            "jsonrpc": "2.0",
            "params": {"db": DB, "login": USERNAME, "password": PASSWORD},
        },
        timeout=60,
    )
    r.raise_for_status()
    data = r.json()
    if "error" in data:
        raise RuntimeError(json.dumps(data["error"], indent=2))
    uid = (data.get("result") or {}).get("uid")
    if not uid:
        raise RuntimeError("Authentication failed (no uid in response)")
    log.info(f"Authenticated as uid={uid}")
    return uid


def get_payment_terms() -> list[str]:
    ctx = {
        "lang": "en_US",
        "tz": "Asia/Dhaka",
        "allowed_company_ids": ALLOWED_COMPANY_IDS,
    }
    res = jsonrpc(
        "fg.stock.dashboard",
        "get_dashboard_data",
        [DASHBOARD_FILTER],
        {"context": ctx},
    )
    terms = [row["Payment Terms"] for row in res.get("payment_terms", [])]
    log.info(f"Found {len(terms)} payment-term buckets on dashboard")
    return terms


def get_drilldown_oas(term: str) -> list[dict]:
    ctx = {
        "lang": "en_US",
        "tz": "Asia/Dhaka",
        "allowed_company_ids": ALLOWED_COMPANY_IDS,
    }
    return jsonrpc(
        "fg.stock.dashboard",
        "get_drilldown_data",
        [DASHBOARD_FILTER, "payment_terms", term, METRIC],
        {"context": ctx},
    ) or []


def get_gspread_client():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for sa_path in [os.path.join(script_dir, "service_account.json"), "service_account.json"]:
        if os.path.exists(sa_path):
            return gspread.service_account(filename=sa_path)
    scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_json = os.getenv("GOOGLE_SHEET_CRED_JSON")
    if creds_json:
        creds_dict = json.loads(creds_json)
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    creds_raw = os.getenv("GOOGLE_CREDS_BASE64")
    if creds_raw:
        try:
            creds_dict = json.loads(creds_raw.strip())
        except json.JSONDecodeError:
            padded = creds_raw.strip() + "=" * (-len(creds_raw.strip()) % 4)
            creds_dict = json.loads(base64.b64decode(padded).decode("utf-8"))
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    raise Exception("No Google credentials found.")


def main():
    if not all([ODOO_URL, DB, USERNAME, PASSWORD]):
        raise SystemExit("Missing ODOO_URL/ODOO_DB/ODOO_USERNAME/ODOO_PASSWORD in .env")

    authenticate()
    terms = get_payment_terms()

    all_rows = []
    for term in terms:
        rows = get_drilldown_oas(term)
        log.info(f"  {term}: {len(rows)} OAs")
        for row in rows:
            all_rows.append({
                "Payment Term": term,
                "OA": row.get("OA", ""),
                "Category": row.get("Category", ""),
                "Value": row.get("Value", 0),
            })

    df = pd.DataFrame(all_rows, columns=["Payment Term", "OA", "Category", "Value"])
    distinct = df["OA"].nunique()
    log.info(f"Total drilldown rows: {len(df)}; distinct OAs: {distinct}; sum Value: {df['Value'].sum():,.0f}")

    log.info(f"Writing to Google Sheets: {WORKSHEET_NAME}")
    client = get_gspread_client()
    sheet = client.open_by_key(SHEET_KEY)
    try:
        ws = sheet.worksheet(WORKSHEET_NAME)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = sheet.add_worksheet(title=WORKSHEET_NAME, rows=max(len(df) + 10, 100), cols=10)
    set_with_dataframe(ws, df, include_index=False, include_column_header=True, resize=True)
    log.info("Done.")


if __name__ == "__main__":
    main()
