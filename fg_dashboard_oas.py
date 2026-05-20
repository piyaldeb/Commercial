"""
Per-OA aging-bucket VALUE report for the FG Stock Dashboard.

Pulls operation.details rows that match the FG dashboard domain
(next_operation=Delivery, state not in done/closed), buckets each line's
value (qty * final_price) by (today - action_date) into 11 aging buckets,
and writes a flat table to the 'FG store-Agieng Wise' Google Sheets tab.

Output columns:
  OA, Company, 0-5 Days, 6-10 Days, 11-15 Days, 16-20 Days, 21-25 Days,
  26-30 Days, 30-60 Days, 61-90 Days, 91-120 Days, 120-180 Days, 180+ Days,
  Total Value
"""

import os
import sys
import json
import time
import base64
import logging
from datetime import date, datetime

import requests
from requests.exceptions import RequestException
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

COMPANY_LABEL = {1: "Zipper", 3: "Metal Trims"}
ALLOWED = [3, 2, 1]
ACTIVE_COMPANY = 3

BUCKETS = [
    ("0-5 Days",     0,   5),
    ("6-10 Days",    6,   10),
    ("11-15 Days",   11,  15),
    ("16-20 Days",   16,  20),
    ("21-25 Days",   21,  25),
    ("26-30 Days",   26,  30),
    ("30-60 Days",   30,  60),
    ("61-90 Days",   61,  90),
    ("91-120 Days",  91,  120),
    ("120-180 Days", 120, 180),
    ("180+ Days",    181, 10**9),
]

session = requests.Session()
USER_ID = None


def retry_request(method, url, max_retries=3, backoff=3, **kwargs):
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


def login():
    global USER_ID
    payload = {"jsonrpc": "2.0", "params": {"db": DB, "login": USERNAME, "password": PASSWORD}}
    r = retry_request(session.post, f"{ODOO_URL}/web/session/authenticate", json=payload)
    result = r.json().get("result") or {}
    if "uid" not in result:
        raise Exception("Login failed")
    USER_ID = result["uid"]
    log.info(f"Logged in (uid={USER_ID})")


def switch_company(company_id):
    payload = {
        "jsonrpc": "2.0", "method": "call",
        "params": {
            "model": "res.users", "method": "write",
            "args": [[USER_ID], {"company_id": company_id}],
            "kwargs": {"context": {"allowed_company_ids": [company_id], "company_id": company_id}},
        },
    }
    r = retry_request(session.post, f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    if "error" in r.json():
        raise Exception(f"switch_company failed: {r.json()['error']}")


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


def fetch_delivery_ops():
    ctx = {
        "lang": "en_US", "tz": "Asia/Dhaka", "uid": USER_ID,
        "allowed_company_ids": ALLOWED, "bin_size": True,
        "current_company_id": ACTIVE_COMPANY,
    }
    spec = {
        "oa_id":       {"fields": {"display_name": {}}},
        "partner_id":  {"fields": {"display_name": {}}},
        "action_date": {},
        "fg_balance":  {},
        "final_price": {},
        "company_id":  {"fields": {}},
    }
    # Same scope as the FG Stock Dashboard: lines awaiting FG Packing with
    # remaining stock balance, not yet shipped.
    domain = [
        ["next_operation", "=", "FG Packing"],
        ["fg_balance", ">", 0],
        ["state", "in", ["waiting", "partial"]],
    ]
    out = []
    offset = 0
    page = 5000
    while True:
        payload = {
            "jsonrpc": "2.0", "method": "call",
            "params": {
                "model": "operation.details", "method": "web_search_read",
                "args": [],
                "kwargs": {
                    "specification": spec, "offset": offset, "order": "id ASC",
                    "limit": page, "context": ctx, "count_limit": 1000001,
                    "domain": domain,
                },
            },
        }
        r = retry_request(
            session.post,
            f"{ODOO_URL}/web/dataset/call_kw/operation.details/web_search_read",
            json=payload,
        )
        body = r.json()
        if "error" in body:
            raise Exception(f"operation.details fetch failed: {body['error']}")
        recs = body.get("result", {}).get("records", [])
        out.extend(recs)
        log.info(f"operation.details: fetched {len(recs)} (total {len(out)})")
        if len(recs) < page:
            break
        offset += page
    return out


def bucket_for_age(age_days):
    for label, lo, hi in BUCKETS:
        if lo <= age_days <= hi:
            return label
    return BUCKETS[-1][0]


def build_rows(ops):
    today = date.today()
    agg = {}  # (oa_id, company_id) -> dict

    for op in ops:
        oa = op.get("oa_id") or {}
        oa_id = oa.get("id")
        if not oa_id:
            continue
        comp = (op.get("company_id") or {}).get("id")
        customer = (op.get("partner_id") or {}).get("display_name", "")
        ad = op.get("action_date")
        if not ad:
            continue
        try:
            ad_dt = datetime.strptime(ad, "%Y-%m-%d %H:%M:%S").date()
        except Exception:
            try:
                ad_dt = pd.to_datetime(ad).date()
            except Exception:
                continue
        age = max((today - ad_dt).days, 0)
        bucket = bucket_for_age(age)

        try:
            balance = float(op.get("fg_balance") or 0)
        except Exception:
            balance = 0.0
        try:
            price = float(op.get("final_price") or 0)
        except Exception:
            price = 0.0
        value = balance * price

        key = (oa_id, comp)
        if key not in agg:
            agg[key] = {
                "OA": oa.get("display_name") or "",
                "Customer": customer,
                "Total Value": 0.0,
                "Company": COMPANY_LABEL.get(comp, ""),
            }
            for label, _, _ in BUCKETS:
                agg[key][label] = 0.0

        row = agg[key]
        row["Total Value"] += value
        row[bucket] += value

    rows = list(agg.values())
    rows.sort(key=lambda r: (-r["Total Value"], r["OA"]))
    return rows


def push_to_sheet(rows):
    cols = ["OA", "Customer"] + [b[0] for b in BUCKETS] + ["Total Value", "Company"]
    df = pd.DataFrame(rows, columns=cols)
    # Round numeric columns to integers for readability
    numeric_cols = [b[0] for b in BUCKETS] + ["Total Value"]
    for c in numeric_cols:
        df[c] = df[c].round(0)

    client = get_gspread_client()
    sheet = client.open_by_key(SHEET_KEY)
    try:
        ws = sheet.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sheet.add_worksheet(
            title=WORKSHEET_NAME,
            rows=max(1000, len(df) + 50),
            cols=max(20, len(df.columns) + 5),
        )
    ws.clear()
    set_with_dataframe(ws, df)
    log.info(f"Pasted {len(df)} rows to '{WORKSHEET_NAME}'")


def main():
    if not all([ODOO_URL, DB, USERNAME, PASSWORD]):
        raise SystemExit("Missing ODOO_URL/ODOO_DB/ODOO_USERNAME/ODOO_PASSWORD in .env")

    login()
    switch_company(ACTIVE_COMPANY)

    ops = fetch_delivery_ops()
    log.info(f"{len(ops)} delivery operation lines fetched")

    rows = build_rows(ops)
    distinct_oas = len({r["OA"] for r in rows})
    total_value = sum(r["Total Value"] for r in rows)
    log.info(f"Built {len(rows)} (OA, company) rows; distinct OAs={distinct_oas}; total value={total_value:,.0f}")

    if os.getenv("FG_DASHBOARD_DRY_RUN"):
        log.info("\n--- DRY RUN: first 3 rows ---")
        for r in rows[:3]:
            log.info(json.dumps(r, indent=2, default=str))
        return

    push_to_sheet(rows)


if __name__ == "__main__":
    main()
