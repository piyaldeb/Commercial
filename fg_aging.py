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
if sys.stderr.encoding and sys.stderr.encoding.lower() != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8")

load_dotenv()
logging.basicConfig(stream=sys.stdout, level=logging.INFO)
log = logging.getLogger()

ODOO_URL = os.getenv("ODOO_URL")
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
            print(f"Attempt {attempt} failed: {e}")
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
    print(f"Logged in (uid={USER_ID})")


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
    creds_json = os.getenv("GOOGLE_SHEET_CRED_JSON")
    if creds_json:
        creds_dict = json.loads(creds_json)
        scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    creds_raw = os.getenv("GOOGLE_CREDS_BASE64")
    if creds_raw:
        try:
            creds_dict = json.loads(creds_raw.strip())
        except json.JSONDecodeError:
            padded = creds_raw.strip() + "=" * (-len(creds_raw.strip()) % 4)
            creds_dict = json.loads(base64.b64decode(padded).decode("utf-8"))
        scopes = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    raise Exception("No Google credentials found.")


def fetch_delivery_ops():
    """Fetch operation.details with next_operation=Delivery, not done/closed — same
    domain the FG dashboard uses to compute aging."""
    ctx = {
        "lang": "en_US", "tz": "Asia/Dhaka", "uid": USER_ID,
        "allowed_company_ids": ALLOWED, "bin_size": True,
        "current_company_id": ACTIVE_COMPANY,
    }
    spec = {
        "oa_id":           {"fields": {"display_name": {}}},
        "partner_id":      {"fields": {"display_name": {}}},
        "action_date":     {},
        "qty":             {},
        "final_price":     {},
        "company_id":      {"fields": {}},
    }
    domain = [
        "&", "&",
        ["next_operation", "=", "Delivery"],
        ["state", "!=", "done"],
        ["state", "!=", "closed"],
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
        print(f"operation.details: fetched {len(recs)} (total {len(out)})")
        if len(recs) < page:
            break
        offset += page
    return out


def fetch_oa_sales_persons(oa_ids):
    """oa_id on operation.details points to sale.order. Pull user_id (sales person)
    for each OA in one paginated batch."""
    if not oa_ids:
        return {}
    ctx = {
        "lang": "en_US", "tz": "Asia/Dhaka", "uid": USER_ID,
        "allowed_company_ids": ALLOWED, "bin_size": True,
        "current_company_id": ACTIVE_COMPANY,
    }
    spec = {"user_id": {"fields": {"display_name": {}}}}
    out = {}
    ids = list(set(oa_ids))
    CHUNK = 500
    for i in range(0, len(ids), CHUNK):
        chunk = ids[i:i + CHUNK]
        offset = 0
        page = 5000
        while True:
            payload = {
                "jsonrpc": "2.0", "method": "call",
                "params": {
                    "model": "sale.order", "method": "web_search_read",
                    "args": [],
                    "kwargs": {
                        "specification": spec, "offset": offset, "order": "id ASC",
                        "limit": page, "context": ctx, "count_limit": 100001,
                        "domain": [["id", "in", chunk]],
                    },
                },
            }
            r = retry_request(
                session.post,
                f"{ODOO_URL}/web/dataset/call_kw/sale.order/web_search_read",
                json=payload,
            )
            body = r.json()
            if "error" in body:
                raise Exception(f"sale.order fetch failed: {body['error']}")
            recs = body.get("result", {}).get("records", [])
            for rec in recs:
                u = rec.get("user_id") or {}
                out[rec["id"]] = u.get("display_name") or ""
            if len(recs) < page:
                break
            offset += page
    print(f"sale.order sales person lookup: {len(out)} OAs")
    return out


def bucket_for_age(age_days):
    for label, lo, hi in BUCKETS:
        if lo <= age_days <= hi:
            return label
    return BUCKETS[-1][0]


def build_rows(ops, oa_to_sales):
    today = date.today()
    agg = {}  # (oa_id, company_id) -> dict

    for op in ops:
        oa = op.get("oa_id") or {}
        oa_id = oa.get("id")
        if not oa_id:
            continue
        comp = (op.get("company_id") or {}).get("id")
        partner = (op.get("partner_id") or {}).get("display_name", "")
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
        age = (today - ad_dt).days
        if age < 0:
            age = 0
        bucket = bucket_for_age(age)

        try:
            qty = float(op.get("qty") or 0)
        except Exception:
            qty = 0.0
        try:
            price = float(op.get("final_price") or 0)
        except Exception:
            price = 0.0
        value = qty * price

        key = (oa_id, comp)
        if key not in agg:
            agg[key] = {
                "OA": oa.get("display_name") or "",
                "Customer": partner,
                "Sales": oa_to_sales.get(oa_id, ""),
                "Company": COMPANY_LABEL.get(comp, ""),
                "Total Qty": 0.0,
                "Total Value": 0.0,
            }
            for label, _, _ in BUCKETS:
                agg[key][f"{label} Qty"] = 0.0
                agg[key][f"{label} Value"] = 0.0

        row = agg[key]
        row["Total Qty"] += qty
        row["Total Value"] += value
        row[f"{bucket} Qty"] += qty
        row[f"{bucket} Value"] += value

    rows = list(agg.values())
    rows.sort(key=lambda r: (-r["Total Value"], r["OA"]))
    return rows


def push_to_sheet(rows):
    cols = ["OA", "Customer", "Sales", "Company", "Total Qty", "Total Value"]
    for label, _, _ in BUCKETS:
        cols.append(f"{label} Qty")
        cols.append(f"{label} Value")
    df = pd.DataFrame(rows, columns=cols)

    client = get_gspread_client()
    sheet = client.open_by_key(SHEET_KEY)
    try:
        ws = sheet.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sheet.add_worksheet(
            title=WORKSHEET_NAME,
            rows=max(1000, len(df) + 50),
            cols=max(40, len(df.columns) + 5),
        )
    ws.clear()
    set_with_dataframe(ws, df)
    print(f"Pasted {len(df)} rows to '{WORKSHEET_NAME}'")


if __name__ == "__main__":
    login()
    switch_company(ACTIVE_COMPANY)

    ops = fetch_delivery_ops()
    print(f"{len(ops)} delivery operation lines fetched")

    oa_ids = sorted({(op.get("oa_id") or {}).get("id")
                     for op in ops if (op.get("oa_id") or {}).get("id")})
    oa_to_sales = fetch_oa_sales_persons(oa_ids)

    rows = build_rows(ops, oa_to_sales)
    print(f"Built {len(rows)} (OA, company) rows")

    if os.getenv("FG_AGING_DRY_RUN"):
        print("\n--- DRY RUN: first 3 rows ---")
        for r in rows[:3]:
            print(json.dumps(r, indent=2, default=str))
        sys.exit(0)

    try:
        push_to_sheet(rows)
    except Exception as e:
        import traceback
        print(f"Sheet push failed: {e}")
        traceback.print_exc()
