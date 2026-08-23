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
  Production Done    manufacturing.order, closing_date in the month, state not
                     'cancel'.  (running order.har)  The orders finished that
                     day: QTY = done_qty, value = that * final_price.  This is
                     the flow that draws Production Pending down.
                     closing_date, not date_order: manufacturing orders inherit
                     the OA's own date, so a "released to production" series is
                     OA Released again to the unit on every day of the month.
                     (date_order is also computed, so Odoo refuses to filter or
                     group on it - only the stored date_order_ twin works.)
  FG Packing         operation.details, next_operation='FG Packing',
                     state not in (cancel, closed), action_date in the month.
                     Value = qty * final_price.  The goods-INTO-FG-store flow,
                     i.e. the summary's 'packed' stage - the same shape as
                     Dispatch, which is the goods-out flow.
                     Verified: 21 of the month's 28 day/company cells match
                     packed to the cent; the other 7 sit within 0.09% of it, the
                     summary being a cron-refreshed snapshot that lags live data.
  Dispatch           operation.details, next_operation='Delivery',
                     state not in (cancel, closed), action_date in the month.
                     (Fg Delivery.har)  Value = qty * final_price.
                     Verified: row count, qty AND value match dispatched exactly
                     for both companies.

  FG Store Balance   operation.details.retrieve_fg_store_datas, called once per
                     day per company with date_from = date_to = that day.
                     (FG closing.har)  QTY/Value = closing_qty/closing_value -
                     what was still sitting in the store when that day ended.
                     This is a DAY-BY-DAY series, not a snapshot: the closing
                     row is a real historical balance, so the metric carries a
                     number on every day of the month up to today.
                     Verified against the two flows either side of it: the
                     day-on-day MOVE in closing equals FG Packing less Dispatch
                     on 22 of the month's 23 days to the cent, the 23rd being
                     67 out of 20,300. The store fills by packing and empties by
                     dispatch, and this series does exactly that.
                     Verified against the snapshot too: 23 Aug closing
                     653,148.23 for Zipper against the FG Store dashboard's own
                     live call, 653,111.32 - 0.006% apart, the live call dropping
                     a 2022 orphan OA that the closing query still carries.
                     One call per DAY, not one per range - a range reading is
                     not just unstitchable, it is wrong. Checked against the FG
                     dashboard's own 1-22 Aug range: closing QTY agrees exactly,
                     5,771,377 pcs, but its closing VALUE is 1,519.00 lower, and
                     all of the difference is one line. OA032360 / C#3 CE, in on
                     31 Jul, shipped out entirely inside the range and left a
                     -0.02 pcs rounding crumb behind; over a range the method
                     hands that line back as closing_value = MINUS its opening
                     value, -1,519.00, for a crumb worth about nothing. Asked for
                     a single day the same line closes at 0.00. So this series
                     reads a touch HIGHER than the dashboard does, and is right
                     to - any line that empties out inside the window drags the
                     dashboard's range total down by its whole opening value.
  Production Pending manufacturing.order, oa_total_balance > 0, oa_id set,
                     state not in (closed, cancel, hold).  (running order.har)
                     Value = balance_qty * final_price.

Production Pending is the one BALANCE left, not a monthly flow: it is what is
outstanding right now, including orders older than this month, so it is NOT
filtered to the month. Its rows are stamped with the run date so they land on
today's column in a date pivot; the document date is kept in "Doc Date". Odoo
keeps no history of it, so there is no honest way to spread it across the month -
which is why Production Done exists: it is the dated event behind the same stage,
and it is what the day columns are for. On the dashboard the row is blank on
every day but the refresh day, so a row of 0.00 never gets mistaken for missing
data.

operation.details and manufacturing.order carry no salesperson, so Sales Person
(and Sales Team / Region where blank) is filled in from the linked sale.order via
oa_id. The closing rows carry no ids at all - only the OA's name - so they are
joined to sale.order by name instead; every OA in the store matched on the month
this was written. One legacy OA (OA400607, 18 pcs in since 2022) has no customer,
team or salesperson on it, and a blank there would drop it out of the dashboard's
"All" wildcard, so those three fields fall back to "(unassigned)".

TABS WRITTEN

  Raw        the table above, rewritten in full on every run.
  Lists      dropdown sources for the dashboard, all formulas over Raw.
  Dashboard  the six metrics down, the days of the selected month across, driven
             by dropdowns for Month / Company / Sales Team / Sales Person /
             Customer / Measure. Every cell is a SUMIFS over Raw, so a selection
             recalculates in the browser without re-running this script. Built
             once and then left alone so the hourly Raw refresh never resets
             someone's selections; RAW_REBUILD_DASHBOARD=1 rewrites it.

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
DASHBOARD_WORKSHEET = "Dashboard"
LISTS_WORKSHEET = "Lists"
# Stamped into Lists!L1. Bump it whenever the layout or a formula changes and the
# next run replaces the built dashboard instead of leaving the stale one alone.
DASHBOARD_VERSION = "4"

COMPANY_IDS = [1, 3]
COMPANY_NAMES = {1: "Zipper", 3: "Metal Trims"}
# Stand-in for a dashboard criterion the source left empty, so the row still
# falls inside the "All" wildcard instead of quietly dropping out of the totals.
UNASSIGNED = "(unassigned)"
DHAKA = timezone(timedelta(hours=6))

# Metric labels, in the order the user's pivot template lists them.
M_PI = "PI issue"
M_OA = "OA Released"
M_LC = "LC Received"
M_PRODUCTION = "Production Done"
M_PACKING = "FG Packing"
M_DISPATCH = "Dispatch"
M_FG = "FG Store Balance"
M_PROD = "Production Pending"
METRIC_ORDER = [M_PI, M_OA, M_LC, M_PRODUCTION, M_PACKING, M_DISPATCH, M_FG, M_PROD]
# A snapshot rather than a daily flow - it carries one date, the run date.
# FG Store Balance used to be one too; retrieve_fg_store_datas gives it a real
# per-day history, so it is a flow-shaped series now and is not in here.
BALANCE_METRICS = {M_PROD}
# A daily series of balances rather than of flows: it has a real number on every
# day up to the last refresh, but 0.00 on the days after it would read as "the
# store emptied" instead of "not yet", so the dashboard stops the row there.
DAILY_BALANCE_METRICS = {M_FG}

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
    ("PI No",        "id"),
    ("OA No",        "id"),
    ("Invoice No",   "id"),
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


def fetch_operations(next_operation, start, end):
    """FG Packing (into the store) and Delivery (out of it) are the same model,
    the same shape and the same date field - only next_operation differs. Value
    is qty * final_price; final_price is the unit price, price_unit is 0 here."""
    domain = [["next_operation", "=", next_operation],
              ["state", "not in", ["cancel", "closed"]],
              ["company_id", "in", COMPANY_IDS],
              ["action_date", ">=", f"{start} 00:00:00"],
              ["action_date", "<=", f"{end} 23:59:59"]]
    return search_read("operation.details", domain,
                       ["oa_id", "company_id", "team_id", "partner_id", "buyer_id",
                        "fg_categ_type", "delivery_code", "action_date", "date_order",
                        "qty", "final_price", "state"])


def fetch_fg_closing(start, end):
    """FG Store Balance, one day at a time. The method takes a date range and
    hands back opening / received / delivery / disposal / closing per (OA, item);
    asked for a single day, its closing IS that day's closing balance, which is
    the only reading that stitches into a series (see the module docstring).

    Company comes from the argument, not from the rows - they carry no company
    at all, and asked for both at once the method merges a pair of them.

    Rows that close at nothing are dropped: 5% of them, all of it stock that was
    delivered out during the day, and they would add nothing but height."""
    out = []
    day = start
    while day <= end:
        for company_id in COMPANY_IDS:
            recs = call_kw("operation.details", "retrieve_fg_store_datas",
                           [[company_id], str(day), str(day)]) or []
            kept = [r for r in recs
                    if round2(r.get("closing_qty")) or round2(r.get("closing_value"))]
            out += [(day, company_id, r) for r in kept]
        day += timedelta(days=1)
    log.info(f"  operation.details (FG closing): {len(out)} rows over "
             f"{(end - start).days + 1} days")
    return out


def fetch_oa_by_name(names):
    """The closing rows name their OA but carry no id, so the sale.order they
    need for Sales Person / Team / Customer has to be looked up by name."""
    names = sorted({clean_text(n) for n in names if n})
    lookup = {}
    for i in range(0, len(names), 500):
        chunk = names[i:i + 500]
        for rec in call_kw("sale.order", "search_read",
                           [[["name", "in", chunk]],
                            ["name", "user_id", "team_id", "region_id",
                             "partner_id", "buyer_name", "lc_number",
                             "order_ref"]], limit=0) or []:
            lookup[clean_text(rec["name"])] = rec
    log.info(f"  sale.order by name: {len(lookup)} of {len(names)} OAs matched")
    return lookup


MO_FIELDS = ["oa_id", "company_id", "partner_id", "buyer_name", "team_id",
             "date_order", "closing_date", "fg_categ_type", "product_template_id",
             "product_uom_qty", "done_qty", "balance_qty", "final_price", "state"]


def fetch_production(start, end):
    """Production Done - manufacturing orders closed off in the month, which is
    production actually completed. done_qty equals the ordered qty on every one
    of them, so nothing is lost by taking the closing day as the day."""
    domain = [["state", "!=", "cancel"],
              ["company_id", "in", COMPANY_IDS],
              ["closing_date", ">=", f"{start} 00:00:00"],
              ["closing_date", "<=", f"{end} 23:59:59"]]
    return search_read("manufacturing.order", domain, MO_FIELDS)


def fetch_running_orders():
    """Production Pending - the running-order list, balance still to produce."""
    domain = ["&", "&", ["oa_total_balance", ">", 0], ["oa_id", "!=", False],
              ["state", "not in", ["closed", "cancel", "hold"]]]
    recs = search_read("manufacturing.order", domain, MO_FIELDS)
    return [r for r in recs if m2o_id(r.get("company_id")) in COMPANY_IDS]


def fetch_released_oas(pi_ids):
    """A PI carries no pointer to its OA - the link is the OA's order_ref, so it
    has to be walked backwards. A PI can be released in several goes (19 of this
    month's 391, up to 8 OAs on one), hence a list per PI."""
    pi_ids = sorted({i for i in pi_ids if i})
    by_pi = {}
    for i in range(0, len(pi_ids), 500):
        chunk = pi_ids[i:i + 500]
        for rec in call_kw("sale.order", "search_read",
                           [[["sales_type", "=", "oa"], ["order_ref", "in", chunk]],
                            ["name", "order_ref"]], limit=0) or []:
            by_pi.setdefault(m2o_id(rec["order_ref"]), []).append(clean_text(rec["name"]))
    log.info(f"  released OAs: {sum(len(v) for v in by_pi.values())} for {len(by_pi)} PIs")
    return by_pi


def fetch_oa_lookup(oa_ids):
    """operation.details and manufacturing.order carry no salesperson; their
    oa_id points at sale.order, which does."""
    oa_ids = sorted({i for i in oa_ids if i})
    lookup = {}
    for i in range(0, len(oa_ids), 500):
        chunk = oa_ids[i:i + 500]
        for rec in call_kw("sale.order", "read",
                           [chunk, ["name", "user_id", "team_id", "region_id",
                                    "partner_id", "buyer_name", "lc_number",
                                    "order_ref"]]) or []:
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


def rows_sale_orders(recs, metric, released_oas=None):
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
        if metric == M_OA:
            # An OA points back at the PI it was released from.
            row["PI No"] = m2o_name(r.get("order_ref"))
            row["OA No"] = clean_text(r.get("name"))
        else:
            # A PI may have been released as several OAs, or none yet.
            row["PI No"] = clean_text(r.get("name"))
            row["OA No"] = ", ".join((released_oas or {}).get(r["id"], []))
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
        row["PI No"] = m2o_name(r.get("pi_number")) or m2o_name(oa.get("order_ref"))
        row["OA No"] = m2o_name(r.get("sale_order"))
        row["Invoice No"] = clean_text(head.get("name")) or m2o_name(r.get("invoice_id"))
        row["LC No"] = clean_text(head.get("lc_no"))
        row["Item"] = clean_text(r.get("fg_categ_type"))
        row["Doc Date"] = clean_date(r.get("invoice_date"))
        row["QTY"] = round2(r.get("quantity"))
        row["Value"] = round2(r.get("price_subtotal"))
        row["Status"] = clean_text(r.get("parent_state"))
        out.append(row)
    return out


def rows_operations(recs, metric, oa_lookup):
    out = []
    for r in recs:
        oa = oa_lookup.get(m2o_id(r.get("oa_id"))) or {}
        row = blank_row()
        stamp(row, metric, r.get("action_date"))
        row["Company"] = m2o_name(r.get("company_id"))
        row["Sales Team"] = m2o_name(r.get("team_id")) or m2o_name(oa.get("team_id"))
        row["Sales Person"] = m2o_name(oa.get("user_id"))
        row["Region"] = m2o_name(oa.get("region_id"))
        row["Customer"] = m2o_name(r.get("partner_id"))
        row["Buyer"] = m2o_name(r.get("buyer_id"))
        row["PI No"] = m2o_name(oa.get("order_ref"))
        row["OA No"] = m2o_name(r.get("oa_id"))
        row["Invoice No"] = clean_text(r.get("delivery_code"))
        row["LC No"] = clean_text(oa.get("lc_number"))
        row["Item"] = clean_text(r.get("fg_categ_type"))
        row["Doc Date"] = clean_date(r.get("date_order"))
        row["QTY"] = round2(r.get("qty"))
        row["Value"] = round2((r.get("qty") or 0) * (r.get("final_price") or 0))
        row["Status"] = clean_text(r.get("state"))
        out.append(row)
    return out


def rows_fg_closing(daily, oa_lookup):
    """Stock still in the FG store at the end of each day of the month."""
    out = []
    for day, company_id, r in daily:
        oa = oa_lookup.get(clean_text(r.get("oa"))) or {}
        row = blank_row()
        stamp(row, M_FG, day)
        row["Company"] = COMPANY_NAMES.get(company_id, "")
        # The three dashboard criteria that must never be blank - one 2022 OA
        # has none of them, and a blank would fall outside the "All" wildcard.
        row["Sales Team"] = m2o_name(oa.get("team_id")) or UNASSIGNED
        row["Sales Person"] = m2o_name(oa.get("user_id")) or UNASSIGNED
        row["Customer"] = m2o_name(oa.get("partner_id")) or UNASSIGNED
        row["Region"] = m2o_name(oa.get("region_id"))
        row["Buyer"] = m2o_name(oa.get("buyer_name"))
        row["PI No"] = m2o_name(oa.get("order_ref"))
        row["OA No"] = clean_text(r.get("oa"))
        row["LC No"] = clean_text(oa.get("lc_number"))
        row["Item"] = clean_text(r.get("item"))
        row["Doc Date"] = clean_date(r.get("in_date"))
        row["QTY"] = round2(r.get("closing_qty"))
        row["Value"] = round2(r.get("closing_value"))
        row["Status"] = clean_text(r.get("lc_status"))
        out.append(row)
    return out


def rows_manufacturing(recs, oa_lookup, metric, qty_field, day_field=None,
                       as_of=None):
    """Production Done is a flow, dated on the day the order closed; Production
    Pending is the balance still to make, so it carries the run date instead."""
    out = []
    for r in recs:
        oa = oa_lookup.get(m2o_id(r.get("oa_id"))) or {}
        row = blank_row()
        stamp(row, metric, as_of or r.get(day_field))
        row["Company"] = m2o_name(r.get("company_id"))
        row["Sales Team"] = m2o_name(r.get("team_id")) or m2o_name(oa.get("team_id"))
        row["Sales Person"] = m2o_name(oa.get("user_id"))
        row["Region"] = m2o_name(oa.get("region_id"))
        row["Customer"] = m2o_name(r.get("partner_id"))
        row["Buyer"] = m2o_name(r.get("buyer_name"))
        row["PI No"] = m2o_name(oa.get("order_ref"))
        row["OA No"] = m2o_name(r.get("oa_id"))
        row["LC No"] = clean_text(oa.get("lc_number"))
        row["Item"] = clean_text(r.get("fg_categ_type")) or m2o_name(r.get("product_template_id"))
        row["Doc Date"] = clean_date(r.get("date_order"))
        row["QTY"] = round2(r.get(qty_field))
        row["Value"] = round2((r.get(qty_field) or 0) * (r.get("final_price") or 0))
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

    log.info("Fetching Production Done ...")
    production = fetch_production(start, end)

    log.info("Fetching FG Packing ...")
    packings = fetch_operations("FG Packing", start, end)

    log.info("Fetching Dispatch ...")
    deliveries = fetch_operations("Delivery", start, end)

    log.info("Fetching FG Store Balance, day by day ...")
    # Only up to today: the closing balance for a day that has not happened yet
    # would just repeat today's and put a flat line across the rest of the month.
    fg_daily = fetch_fg_closing(start, min(end, as_of))
    fg_oas = fetch_oa_by_name([r.get("oa") for _, _, r in fg_daily])

    log.info("Fetching Production Pending ...")
    running = fetch_running_orders()

    oa_lookup = fetch_oa_lookup([m2o_id(r.get("oa_id")) for r in packings]
                                + [m2o_id(r.get("oa_id")) for r in deliveries]
                                + [m2o_id(r.get("oa_id")) for r in production]
                                + [m2o_id(r.get("oa_id")) for r in running]
                                + [m2o_id(r.get("sale_order")) for r in invoice_lines])

    released_oas = fetch_released_oas([r["id"] for r in pis])

    rows = (rows_sale_orders(pis, M_PI, released_oas)
            + rows_sale_orders(oas, M_OA)
            + rows_invoice_lines(invoice_lines, headers, oa_lookup)
            + rows_manufacturing(production, oa_lookup, M_PRODUCTION, "done_qty",
                                 day_field="closing_date")
            + rows_operations(packings, M_PACKING, oa_lookup)
            + rows_operations(deliveries, M_DISPATCH, oa_lookup)
            + rows_fg_closing(fg_daily, fg_oas)
            + rows_manufacturing(running, oa_lookup, M_PROD, "balance_qty", as_of=as_of))

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
PALE = {"red": 0.85, "green": 0.89, "blue": 0.96}
PALE_YELLOW = {"red": 1.0, "green": 0.95, "blue": 0.80}


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



# ========= DASHBOARD ==========
# A formula-driven dashboard rather than a pivot table: the selectors are plain
# data-validation dropdowns and every cell is a SUMIFS over "Raw", so changing a
# selection recalculates in the browser - no re-run of this script needed. It is
# built ONCE and then left alone, so the hourly Raw refresh never clobbers the
# selections someone left sitting in it. RAW_REBUILD_DASHBOARD=1 rewrites it.
#
# Raw column letters the formulas depend on:
#   A Metric   B Date   C Month   D Company   E Sales Team
#   F Sales Person      G Region  H Customer  P QTY   Q Value
#
# Selector cells on Dashboard:
#   B3 Month   B4 Company   B5 Sales Team   B6 Sales Person   B7 Customer
#   B8 Measure (QTY / Value)
#
# The Sales Person and Customer lists are DEPENDENT - they narrow to whichever
# Sales Team is selected, so picking JAMUNA leaves only JAMUNA's people and
# customers in the next two dropdowns. Validation is deliberately non-strict:
# changing the team can leave a stale name behind in B6/B7, and a warning
# triangle is friendlier than a hard rejection.
#
# "All" is the "*" SUMIFS wildcard, which matches any non-empty text. Safe here
# because Company / Sales Team / Sales Person / Customer are never blank on any
# row - checked across all six metrics.

FIRST_DAY_COL = 2                      # B
LAST_DAY_COL = FIRST_DAY_COL + 30      # AF, 31 days
TOTAL_COL = LAST_DAY_COL + 1           # AG
HEADER_ROW = 10
FIRST_METRIC_ROW = HEADER_ROW + 1

# selector row, label, dropdown source range, default
SELECTORS = [
    (3, "Month",        "Lists!$G$1:$G$40",  None),
    (4, "Company",      "Lists!$I$1:$I$20",  "All"),
    (5, "Sales Team",   "Lists!$A$1:$A$60",  "All"),
    (6, "Sales Person", "Lists!$C$1:$C$120", "All"),
    (7, "Customer",     "Lists!$E$1:$E$800", "All"),
]


def a1_col(index):
    """1-based column number -> column letters."""
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def lists_formulas():
    """Dropdown sources. Sales Person and Customer narrow to the selected team."""
    team = "Dashboard!$B$5"
    person = "Dashboard!$B$6"
    return [
        ("A1", "All"),
        ("A2", '=IFERROR(SORT(UNIQUE(FILTER(Raw!$E$2:$E,Raw!$E$2:$E<>""))),"")'),
        ("C1", "All"),
        ("C2", '=IFERROR(SORT(UNIQUE(FILTER(Raw!$F$2:$F,Raw!$F$2:$F<>"",'
               'IF(' + team + '="All",Raw!$F$2:$F<>"",Raw!$E$2:$E=' + team + ')))),"")'),
        ("E1", "All"),
        ("E2", '=IFERROR(SORT(UNIQUE(FILTER(Raw!$H$2:$H,Raw!$H$2:$H<>"",'
               'IF(' + team + '="All",Raw!$H$2:$H<>"",Raw!$E$2:$E=' + team + '),'
               'IF(' + person + '="All",Raw!$H$2:$H<>"",Raw!$F$2:$F=' + person + ')))),"")'),
        ("G1", '=IFERROR(SORT(UNIQUE(FILTER(Raw!$C$2:$C,Raw!$C$2:$C<>""))),"")'),
        ("I1", "All"),
        ("I2", '=IFERROR(SORT(UNIQUE(FILTER(Raw!$D$2:$D,Raw!$D$2:$D<>""))),"")'),
        # Which Raw column the Measure dropdown points at, for the SUMIFS below.
        ("K1", '=IF(Dashboard!$B$8="QTY","P","Q")'),
        ("L1", DASHBOARD_VERSION),
    ]


def dashboard_formulas(default_month):
    cells = [("A1", "Commercial Pipeline Dashboard")]
    for row, label, _, default in SELECTORS:
        cells.append(("A" + str(row), label))
        if default is not None:
            cells.append(("B" + str(row), default))
    cells += [("B3", default_month), ("A8", "Measure"), ("B8", "Value")]

    # Day header: the 1st of the selected month, then +1 for as long as the
    # result is still inside that month, so short months leave blanks.
    # B3 is meant to hold the text "2026-08", but Sheets parses that shape as a
    # date on entry, so the cell can legitimately end up either. ISNUMBER tells
    # them apart; EOMONTH(x,-1)+1 walks a date back to the 1st of its own month.
    month_start = ('IF(ISNUMBER($B$3),EOMONTH($B$3,-1)+1,'
                   'DATEVALUE($B$3&"-01"))')
    first = a1_col(FIRST_DAY_COL) + str(HEADER_ROW)
    anchor = "$" + a1_col(FIRST_DAY_COL) + "$" + str(HEADER_ROW)
    cells.append((first, '=IF($B$3="","",' + month_start + ')'))
    for col in range(FIRST_DAY_COL + 1, LAST_DAY_COL + 1):
        prev = a1_col(col - 1) + str(HEADER_ROW)
        cells.append((a1_col(col) + str(HEADER_ROW),
                      '=IF(OR($B$3="",' + prev + '=""),"",'
                      'IF(MONTH(' + prev + '+1)=MONTH(' + anchor + '),'
                      + prev + '+1,""))'))
    cells.append(("A" + str(HEADER_ROW), "Details"))
    cells.append((a1_col(TOTAL_COL) + str(HEADER_ROW), "Total"))

    measure = 'INDIRECT("Raw!$"&Lists!$K$1&"$2:$"&Lists!$K$1)'
    criteria = [("$D$2:$D", "$B$4"), ("$E$2:$E", "$B$5"),
                ("$F$2:$F", "$B$6"), ("$H$2:$H", "$B$7")]
    for offset, metric in enumerate(METRIC_ORDER):
        row = FIRST_METRIC_ROW + offset
        cells.append(("A" + str(row), metric))
        for col in range(FIRST_DAY_COL, LAST_DAY_COL + 1):
            head = a1_col(col) + "$" + str(HEADER_ROW)
            parts = [measure, "Raw!$A$2:$A,$A" + str(row), "Raw!$B$2:$B," + head]
            parts += ['Raw!' + rng + ',IF(' + sel + '="All","*",' + sel + ')'
                      for rng, sel in criteria]
            body = 'SUMIFS(' + ",".join(parts) + ')'
            if metric in DAILY_BALANCE_METRICS:
                # Blank the days past the last one Raw holds - a flow is
                # honestly 0 before it happens, a balance is simply unknown.
                body = ('IF(' + head + '>MAXIFS(Raw!$B$2:$B,Raw!$A$2:$A,$A'
                        + str(row) + '),"",' + body + ')')
            if metric in BALANCE_METRICS:
                # A balance carries a single date - the last refresh - so every
                # other day would read 0.00 and be taken for missing data. Blank
                # them, and let MAXIFS find the refresh day so the formula keeps
                # up with the hourly run on its own.
                body = ('IF(' + head + '<>MAXIFS(Raw!$B$2:$B,Raw!$A$2:$A,$A'
                        + str(row) + '),"",' + body + ')')
            cells.append((a1_col(col) + str(row),
                          '=IF(' + head + '="","",' + body + ')'))
        span = (a1_col(FIRST_DAY_COL) + str(row) + ':'
                + a1_col(LAST_DAY_COL) + str(row))
        if metric in DAILY_BALANCE_METRICS:
            # Adding up 31 closing balances would count the same stock 31 times.
            # The month's total for a balance is the last day that has one -
            # LOOKUP(2,1/...) is the idiom for "last non-blank in the row".
            total = 'IFERROR(LOOKUP(2,1/(' + span + '<>""),' + span + '),"")'
        else:
            total = 'SUM(' + span + ')'
        cells.append((a1_col(TOTAL_COL) + str(row),
                      '=IF($B$3="","",' + total + ')'))

    note_row = FIRST_METRIC_ROW + len(METRIC_ORDER) + 1
    cells.append(("A" + str(note_row),
                  "The first six rows are daily flows - what happened on each "
                  "day. FG Store Balance is a daily closing balance: what was "
                  "still in the store when that day ended, so it does not add up "
                  "across days and its Total column shows the last day's "
                  "closing instead of a sum. Production Pending "
                  "is a balance as of the last refresh, so it carries a number "
                  "on the refresh day only and stays blank on every other day; "
                  "for the day-by-day version of that stage read Production Done "
                  "(orders finished in the factory)."))
    return cells


def ensure_dashboard(sheet, default_month):
    existing = {ws.title for ws in sheets_call(sheet.worksheets)}
    rebuild = bool(os.getenv("RAW_REBUILD_DASHBOARD"))
    if DASHBOARD_WORKSHEET in existing and not rebuild:
        # Rebuilding every hour would wipe whatever the user had selected, so it
        # only happens when the built layout is older than the code.
        built = ""
        if LISTS_WORKSHEET in existing:
            try:
                built = str(sheets_call(
                    sheet.worksheet(LISTS_WORKSHEET).acell, "L1").value or "")
            except Exception:
                built = ""
        if built.strip() == DASHBOARD_VERSION:
            log.info(f"'{DASHBOARD_WORKSHEET}' is at v{DASHBOARD_VERSION}, "
                     f"leaving it alone")
            return
        log.info(f"'{DASHBOARD_WORKSHEET}' is at v{built or '?'}, rebuilding at "
                 f"v{DASHBOARD_VERSION} (dropdown selections reset)")

    # Both tabs have to exist before either is written: the Lists formulas
    # reference Dashboard! and the Dashboard formulas reference Lists!, and a
    # formula naming a sheet that does not exist yet sticks as #REF! rather than
    # resolving once the sheet appears.
    if LISTS_WORKSHEET in existing:
        lists_ws = sheet.worksheet(LISTS_WORKSHEET)
    else:
        lists_ws = sheets_call(sheet.add_worksheet, title=LISTS_WORKSHEET,
                               rows=1000, cols=12)
    if DASHBOARD_WORKSHEET in existing:
        ws = sheet.worksheet(DASHBOARD_WORKSHEET)
    else:
        ws = sheets_call(sheet.add_worksheet, title=DASHBOARD_WORKSHEET,
                         rows=40, cols=TOTAL_COL + 2)
    if ws.row_count < 40 or ws.col_count < TOTAL_COL + 2:
        sheets_call(ws.resize, rows=max(ws.row_count, 40),
                    cols=max(ws.col_count, TOTAL_COL + 2))

    sheets_call(lists_ws.clear)
    sheets_call(ws.clear)
    sheets_call(lists_ws.batch_update,
                [{"range": ref, "values": [[val]]} for ref, val in lists_formulas()],
                value_input_option="USER_ENTERED")
    # TEXT on B3 before the write, or USER_ENTERED turns "2026-08" into the date
    # serial for 1 Aug. It still *displays* as 2026-08, so the damage is invisible
    # until $B$3&"-01" builds "46235-01" and the whole grid goes #VALUE!.
    sheets_call(sheet.batch_update, {"requests": [{"repeatCell": {
        "range": {"sheetId": ws.id, "startRowIndex": 2, "endRowIndex": 3,
                  "startColumnIndex": 1, "endColumnIndex": 2},
        "cell": {"userEnteredFormat": {"numberFormat": {"type": "TEXT"}}},
        "fields": "userEnteredFormat.numberFormat"}}]})
    sheets_call(ws.batch_update,
                [{"range": ref, "values": [[val]]}
                 for ref, val in dashboard_formulas(default_month)],
                value_input_option="USER_ENTERED")

    sid = ws.id
    last_metric_row = FIRST_METRIC_ROW + len(METRIC_ORDER)
    balance_rows = [FIRST_METRIC_ROW + i for i, m in enumerate(METRIC_ORDER)
                    if m in BALANCE_METRICS or m in DAILY_BALANCE_METRICS]
    reqs = []
    for row, _, source, _ in SELECTORS:
        reqs.append({"setDataValidation": {
            "range": {"sheetId": sid, "startRowIndex": row - 1, "endRowIndex": row,
                      "startColumnIndex": 1, "endColumnIndex": 2},
            "rule": {"condition": {"type": "ONE_OF_RANGE",
                                   "values": [{"userEnteredValue": "=" + source}]},
                     "showCustomUi": True, "strict": False}}})
    reqs.append({"setDataValidation": {
        "range": {"sheetId": sid, "startRowIndex": 7, "endRowIndex": 8,
                  "startColumnIndex": 1, "endColumnIndex": 2},
        "rule": {"condition": {"type": "ONE_OF_LIST",
                               "values": [{"userEnteredValue": "QTY"},
                                          {"userEnteredValue": "Value"}]},
                 "showCustomUi": True, "strict": True}}})
    reqs += [
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": 4},
            "cell": {"userEnteredFormat": {
                "textFormat": {"bold": True, "fontSize": 14}}},
            "fields": "userEnteredFormat.textFormat"}},
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 2, "endRowIndex": 8,
                      "startColumnIndex": 0, "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
            "fields": "userEnteredFormat.textFormat"}},
        # Shade the selector cells, so it is obvious what you are meant to click.
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 2, "endRowIndex": 8,
                      "startColumnIndex": 1, "endColumnIndex": 2},
            "cell": {"userEnteredFormat": {"backgroundColor": PALE_YELLOW}},
            "fields": "userEnteredFormat.backgroundColor"}},
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": HEADER_ROW - 1,
                      "endRowIndex": HEADER_ROW, "startColumnIndex": 0,
                      "endColumnIndex": TOTAL_COL},
            "cell": {"userEnteredFormat": {
                "backgroundColor": BLUE,
                "horizontalAlignment": "CENTER",
                "numberFormat": {"type": "DATE", "pattern": "d-mmm"},
                "textFormat": {"bold": True, "foregroundColor": {
                    "red": 1, "green": 1, "blue": 1}}}},
            "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,"
                      "numberFormat,textFormat)"}},
        # "Details" and "Total" are labels, not dates - undo the date format.
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": HEADER_ROW - 1,
                      "endRowIndex": HEADER_ROW, "startColumnIndex": 0,
                      "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "TEXT"}}},
            "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": HEADER_ROW - 1,
                      "endRowIndex": HEADER_ROW,
                      "startColumnIndex": TOTAL_COL - 1, "endColumnIndex": TOTAL_COL},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "TEXT"}}},
            "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": HEADER_ROW,
                      "endRowIndex": last_metric_row, "startColumnIndex": 0,
                      "endColumnIndex": 1},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}},
            "fields": "userEnteredFormat.textFormat"}},
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": HEADER_ROW,
                      "endRowIndex": last_metric_row, "startColumnIndex": 1,
                      "endColumnIndex": TOTAL_COL},
            "cell": {"userEnteredFormat": {
                "numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}}},
            "fields": "userEnteredFormat.numberFormat"}},
        # Italic on the balance rows: they are read differently from the flows
        # above them - neither one adds up across the days, and Production
        # Pending is mostly blank besides.
        {"repeatCell": {
            "range": {"sheetId": sid,
                      "startRowIndex": min(balance_rows) - 1,
                      "endRowIndex": max(balance_rows),
                      "startColumnIndex": 0, "endColumnIndex": TOTAL_COL},
            "cell": {"userEnteredFormat": {"textFormat": {"italic": True}}},
            "fields": "userEnteredFormat.textFormat.italic"}},
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": HEADER_ROW,
                      "endRowIndex": last_metric_row,
                      "startColumnIndex": TOTAL_COL - 1, "endColumnIndex": TOTAL_COL},
            "cell": {"userEnteredFormat": {"textFormat": {"bold": True},
                                           "backgroundColor": PALE}},
            "fields": "userEnteredFormat(textFormat,backgroundColor)"}},
        {"updateSheetProperties": {
            "properties": {"sheetId": sid, "gridProperties": {
                "frozenRowCount": HEADER_ROW, "frozenColumnCount": 1}},
            "fields": "gridProperties(frozenRowCount,frozenColumnCount)"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": 0, "endIndex": 1},
            "properties": {"pixelSize": 165}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": 1, "endIndex": TOTAL_COL},
            "properties": {"pixelSize": 78}, "fields": "pixelSize"}},
    ]
    sheets_call(sheet.batch_update, {"requests": reqs})
    log.info(f"Built '{DASHBOARD_WORKSHEET}' and '{LISTS_WORKSHEET}'")


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
    ensure_dashboard(sheet, as_of.strftime("%Y-%m"))
    log.info("Done")


if __name__ == "__main__":
    main()
