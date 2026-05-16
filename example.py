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

from datetime import date
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

SHEET_KEY = "16KUtHe-6R6Sd1HarFgS6Y2rEZGmUcM3ohAzPLy7Kxkg"
WORKSHEET_NAME = "PO_Details"

today = date.today()

session = requests.Session()
USER_ID = None

# ========= MAPPINGS ==========
STATE_MAP = {
    "draft":      "Draft",
    "sent":       "RFQ Sent",
    "to approve": "To Approve",
    "purchase":   "Approved",
    "done":       "Locked",
    "cancel":     "Cancelled",
}

# ========= COLUMN ORDER ==========
COLUMNS = [
    "Company", "PO No", "PO Apprvd Stat", "P Cat", "P Type",
    "Inv Month", "Vendor", "Item Details", "Odoo Code", "Inv No", "Inv Date",
    "Inv Quantity", "Inv Value", "Adjust", "Pmt Term", "Ship Mode",
    "Inco", "Booked Ship ETD", "Booked Ship ETA", "ETD", "ETA",
    "BL Number", "BL Date", "LC Number", "LC Date",
    "I/H Plan Month", "Inhoused Date", "I/H Status",
]

# ========= GOOGLE SHEETS CLIENT ==========
def get_gspread_client():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for sa_path in [
        os.path.join(script_dir, "service_account.json"),
        "service_account.json",
    ]:
        if os.path.exists(sa_path):
            return gspread.service_account(filename=sa_path)

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

# ========= FETCH TRANSIT LINES ==========
def fetch_po_details(company_id, cname):
    ctx = {
        "lang": "en_US", "tz": "Asia/Dhaka", "uid": USER_ID,
        "allowed_company_ids": [company_id], "bin_size": True,
        "current_company_id": company_id,
    }

    # Step 1: Fetch transit.line records (po_id only display_name — deeper fields blocked by ACL)
    line_spec = {
        "transit_id":     {},
        "po_id":          {"fields": {"display_name": {}}},
        "invoice_number": {},
        "invoice_date":   {},
        "vendor_main":    {"fields": {"display_name": {}}},
        "product_id":     {"fields": {"display_name": {}}},
        "qty_in_transit": {},
        "subtotal":       {},
        "shipment_mode":  {"fields": {"display_name": {}}},
        "bl_number":      {},
        "lc_number":      {},
        "eta":            {},
        "state":          {},
    }
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "transit.line",
            "method": "web_search_read",
            "args": [],
            "kwargs": {
                "specification": line_spec,
                "offset": 0,
                "order": "invoice_date DESC",
                "limit": 10000,
                "context": ctx,
                "count_limit": 100001,
                "domain": [["transit_id.company_id", "=", company_id]],
            },
        },
    }
    r = retry_request(
        session.post,
        f"{ODOO_URL}/web/dataset/call_kw/transit.line/web_search_read",
        json=payload,
    )
    r.raise_for_status()

    try:
        data = r.json()["result"]["records"]
    except Exception as e:
        print(f"❌ {cname}: Failed to parse response: {e}")
        print(r.text[:2000])
        return []

    print(f"📋 {cname}: {len(data)} transit lines fetched")

    # Step 2: Fetch purchase.order records for PO Apprvd Stat
    po_ids = list({(rec.get("po_id") or {}).get("id") for rec in data if (rec.get("po_id") or {}).get("id")})
    po_map = {}
    if po_ids:
        po_payload = {
            "jsonrpc": "2.0",
            "method": "call",
            "params": {
                "model": "purchase.order",
                "method": "web_search_read",
                "args": [],
                "kwargs": {
                    "specification": {
                        "name":           {},
                        "display_name":   {},
                        "state":          {},
                        "next_approver":  {"fields": {"display_name": {}}},
                    },
                    "offset": 0,
                    "order": "id ASC",
                    "limit": len(po_ids) + 10,
                    "context": ctx,
                    "count_limit": 100001,
                    "domain": [["id", "in", po_ids]],
                },
            },
        }
        pr = retry_request(session.post, f"{ODOO_URL}/web/dataset/call_kw/purchase.order/web_search_read", json=po_payload)
        for p in pr.json().get("result", {}).get("records", []):
            po_map[p["id"]] = p
        print(f"📋 {cname}: {len(po_map)} purchase orders fetched")

    # Step 2b: Fetch product.template for P Cat (categ_type) and P Type (classification_id) by Odoo code
    def _extract_code(rec):
        dn = (rec.get("product_id") or {}).get("display_name", "") or ""
        if dn.startswith("[") and "]" in dn:
            return dn[1:dn.index("]")]
        return ""

    all_codes = list({_extract_code(rec) for rec in data if _extract_code(rec)})
    product_map = {}
    chunk_size = 500
    for i in range(0, len(all_codes), chunk_size):
        chunk = all_codes[i:i + chunk_size]
        prod_payload = {
            "jsonrpc": "2.0",
            "method": "call",
            "params": {
                "model": "product.template",
                "method": "web_search_read",
                "args": [],
                "kwargs": {
                    "specification": {
                        "default_code":      {},
                        "categ_type":        {"fields": {"display_name": {}}},
                        "classification_id": {"fields": {"display_name": {}}},
                    },
                    "offset": 0,
                    "order": "id ASC",
                    "limit": len(chunk) + 10,
                    "context": ctx,
                    "count_limit": 100001,
                    "domain": [["default_code", "in", chunk]],
                },
            },
        }
        pr2 = retry_request(session.post, f"{ODOO_URL}/web/dataset/call_kw/product.template/web_search_read", json=prod_payload)
        for rec in pr2.json().get("result", {}).get("records", []):
            code = rec.get("default_code") or ""
            if code:
                product_map[code] = {
                    "P Cat":  (rec.get("categ_type") or {}).get("display_name", ""),
                    "P Type": (rec.get("classification_id") or {}).get("display_name", ""),
                }
    print(f"📦 {cname}: {len(product_map)} products mapped (P Cat / P Type)")

    # Step 3: Fetch transit.model header records for remaining fields
    def _tid(rec):
        t = rec.get("transit_id")
        if isinstance(t, dict):
            return t.get("id")
        return t if t else None

    transit_ids = list({_tid(rec) for rec in data if _tid(rec)})
    transit_map = {}
    if transit_ids:
        hdr_payload = {
            "jsonrpc": "2.0",
            "method": "call",
            "params": {
                "model": "transit.model",
                "method": "web_search_read",
                "args": [],
                "kwargs": {
                    "specification": {
                        "company_id":     {"fields": {"display_name": {}}},
                        "adjusted_state": {},
                        "payment_term":   {"fields": {"display_name": {}}},
                        "inco_terms":     {"fields": {"display_name": {}}},
                        "booked_etd":     {},
                        "booked_eta":     {},
                        "etd":            {},
                        "bl_date":        {},
                        "lc_date":        {},
                        "ih_plan":        {},
                        "grn_date":       {},
                    },
                    "offset": 0,
                    "order": "id ASC",
                    "limit": len(transit_ids) + 10,
                    "context": ctx,
                    "count_limit": 100001,
                    "domain": [["id", "in", transit_ids]],
                },
            },
        }
        hr = retry_request(session.post, f"{ODOO_URL}/web/dataset/call_kw/transit.model/web_search_read", json=hdr_payload)
        for h in hr.json().get("result", {}).get("records", []):
            transit_map[h["id"]] = h
        print(f"📋 {cname}: {len(transit_map)} transit headers fetched")

    def map_record(rec):
        inv_date = rec.get("invoice_date") or ""
        try:
            inv_month = pd.to_datetime(inv_date).strftime("%b-%y") if inv_date else ""
        except Exception:
            inv_month = ""

        po_ref = rec.get("po_id") or {}
        po = po_map.get(po_ref.get("id"), {})
        product = rec.get("product_id") or {}
        tid = _tid(rec)
        transit = transit_map.get(tid, {})

        dn = product.get("display_name", "") or ""
        if dn.startswith("[") and "]" in dn:
            odoo_code = dn[1:dn.index("]")]
            item_details = dn[dn.index("]") + 2:]
        else:
            odoo_code = ""
            item_details = dn

        subtotal = rec.get("subtotal")
        qty = rec.get("qty_in_transit")

        return {
            "Company":         (transit.get("company_id") or {}).get("display_name", ""),
            "PO No":           po.get("name", "") or po_ref.get("display_name", ""),
            "PO Apprvd Stat":  (
                STATE_MAP.get(po.get("state", ""), po.get("state", ""))
                if po.get("state") in ("purchase", "cancel", "done", "")
                else (po.get("next_approver") or {}).get("display_name", "")
                or STATE_MAP.get(po.get("state", ""), po.get("state", ""))
            ),
            "P Cat":           product_map.get(odoo_code, {}).get("P Cat", ""),
            "P Type":          product_map.get(odoo_code, {}).get("P Type", ""),
            "Inv Month":       inv_month,
            "Vendor":          (rec.get("vendor_main") or {}).get("display_name", ""),
            "Item Details":    item_details,
            "Odoo Code":       odoo_code,
            "Inv No":          rec.get("invoice_number") or "",
            "Inv Date":        inv_date,
            "Inv Quantity":    qty if qty is not None else "",
            "Inv Value":       subtotal if subtotal is not None else "",
            "Adjust":          transit.get("adjusted_state") or "",
            "Pmt Term":        (transit.get("payment_term") or {}).get("display_name", ""),
            "Ship Mode":       (rec.get("shipment_mode") or {}).get("display_name", ""),
            "Inco":            (transit.get("inco_terms") or {}).get("display_name", ""),
            "Booked Ship ETD": transit.get("booked_etd") or "",
            "Booked Ship ETA": transit.get("booked_eta") or "",
            "ETD":             transit.get("etd") or "",
            "ETA":             rec.get("eta") or "",
            "BL Number":       rec.get("bl_number") or "",
            "BL Date":         transit.get("bl_date") or "",
            "LC Number":       rec.get("lc_number") or "",
            "LC Date":         transit.get("lc_date") or "",
            "I/H Plan Month":  transit.get("ih_plan") or "",
            "Inhoused Date":   transit.get("grn_date") or "",
            "I/H Status":      rec.get("state") or "",
        }

    all_rows = [map_record(rec) for rec in data]
    print(f"📊 {cname}: {len(all_rows)} transit lines fetched")

    # ---- coverage check ----
    df_check = pd.DataFrame(all_rows, columns=COLUMNS)
    empty = {c: int(df_check[c].replace("", pd.NA).isna().sum())
             for c in COLUMNS if df_check[c].replace("", pd.NA).isna().any()}
    if empty:
        print("⚠️  Columns with blank values:")
        for col, n in empty.items():
            print(f"   {col}: {n} blank row(s)")
    else:
        print("✅ No blank values in any column")

    return all_rows

# ========= MAIN ==========
if __name__ == "__main__":
    userinfo = login()
    print("User info (allowed companies):", userinfo.get("user_companies", {}))

    company_id = 3
    cname = "Metal Trims"

    if switch_company(company_id):
        records = fetch_po_details(company_id, cname)

        if records:
            df = pd.DataFrame(records, columns=COLUMNS)
            output_file = f"po_details_{today.isoformat()}.xlsx"
            df.to_excel(output_file, index=False)
            print(f"📂 Saved: {output_file}")

            # ========= GOOGLE SHEETS ==========
            try:
                client = get_gspread_client()
                sheet = client.open_by_key(SHEET_KEY)
                worksheet = sheet.worksheet(WORKSHEET_NAME)
                worksheet.batch_clear(["A:AB"])
                set_with_dataframe(worksheet, df)
                print(f"✅ Data pasted to Google Sheets → '{WORKSHEET_NAME}'")
            except Exception as e:
                import traceback
                print(f"❌ Error while pasting to Google Sheets: {e}")
                traceback.print_exc()
        else:
            print(f"❌ No data fetched for {cname}")