"""
OYA MICROCREDIT — MONTHLY REVIEW SCRIPT (Python)
=================================================
Run once per month after uploading files to Drive.

SETUP (one-time):
  1. pip install -r requirements.txt
  2. Create a Google Service Account, share your Drive folder and Sheet with it
  3. Download the service account JSON key → save as credentials.json (same folder as this script)
  4. Fill in CONFIG below (Sheet ID, folder ID, GitHub token)
  5. Run: python process.py

USAGE:
  python process.py              # processes prior month automatically
  python process.py 2026-02      # processes a specific month
"""

import sys
import os
import io
import json
import base64
import logging
import re
from datetime import date, datetime, timedelta
from calendar import monthrange

import pandas as pd
import numpy as np
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ─── CONFIG (edit these) ─────────────────────────────────────────────────────

CONFIG = {
    "SHEET_ID":    "1bzOLhOm31PpQ2GR5mNzt5SIcl1xzfoATPI_WKiLN7xI",
    "DRIVE_ROOT":  "1Njvt_y4KeE5-b5kEXJz2X0_uZMgsgTCM",
    "CREDENTIALS": "credentials.json",   # path to your service account key file

    "GITHUB": {
        "OWNER":  "ikyei",
        "REPO":   "oya-monthly",
        "BRANCH": "main",
        # Store your GitHub token in an environment variable called GITHUB_TOKEN
        # (or paste it directly here — but don't commit it to GitHub!)
        "TOKEN":  os.environ.get("GITHUB_TOKEN", ""),
    },

    "COUNTRIES": ["TZ", "KE", "UG", "SL", "NG"],

    "COUNTRY_NAMES": {
        "TZ": "Tanzania",
        "KE": "Kenya",
        "UG": "Uganda",
        "SL": "Sierra Leone",
        "NG": "Nigeria",
    },

    # Fallback column positions if header detection fails (0-based)
    "COLUMN_MAPS": {
        "app_date_index":        12,
        "app_team_index":        14,
        "app_area_index":        15,
        "apr_date_index":        12,
        "disb_date_index":        0,
        "disb_principal_index":   4,
        "disb_team_index":       14,
        "stat_principal_index":   5,
        "stat_interest_index":    6,
        "stat_overdue_age_index": 9,
    },

    "PL_STAFF_KEYS":   ["staff costs", "staff cost", "salaries"],
    "PL_FUEL_KEYS":    ["vehicle fuel", "fuel"],
    "PL_VEHICLE_KEYS": ["vehicle maintenance"],
}

# ─── LOGGING ─────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ─── GOOGLE API SETUP ─────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

def get_services():
    """Authenticate and return (drive_service, sheets_service)."""
    creds = service_account.Credentials.from_service_account_file(
        CONFIG["CREDENTIALS"], scopes=SCOPES
    )
    drive  = build("drive",  "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)
    return drive, sheets

# ─── DRIVE HELPERS ───────────────────────────────────────────────────────────

def list_folder(drive, folder_id):
    """Returns list of {id, name, mimeType, modifiedTime} in a folder."""
    results, page_token = [], None
    while True:
        resp = drive.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="nextPageToken, files(id,name,mimeType,modifiedTime)",
            pageToken=page_token,
            pageSize=200,
        ).execute()
        results.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return results

def find_subfolder(drive, parent_id, name):
    """Find a subfolder by name inside parent_id. Returns folder id or None."""
    items = list_folder(drive, parent_id)
    for item in items:
        if item["mimeType"] == "application/vnd.google-apps.folder" and item["name"] == name:
            return item["id"]
    return None

def get_spreadsheet_files(drive, folder_id):
    """Returns spreadsheet files (.xlsx/.xls/.csv) sorted newest-modified first."""
    items = list_folder(drive, folder_id)
    files = [
        f for f in items
        if f["name"].lower().endswith((".xlsx", ".xls", ".xlt", ".csv"))
        and f["mimeType"] != "application/vnd.google-apps.folder"
    ]
    files.sort(key=lambda f: f.get("modifiedTime", ""), reverse=True)
    return files

def download_file(drive, file_meta) -> pd.DataFrame | dict:
    """
    Downloads a Drive file and returns a dict of {sheet_name: DataFrame}.
    For CSV returns {"Sheet1": DataFrame}.
    For xlsx/xls returns all sheets.
    """
    name = file_meta["name"].lower()
    file_id = file_meta["id"]

    fh = io.BytesIO()
    request = drive.files().get_media(fileId=file_id)
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)

    if name.endswith(".csv") or name.endswith(".txt"):
        try:
            df = pd.read_csv(fh, dtype=str, keep_default_na=False, on_bad_lines='skip', encoding="utf-8")
            return {"Sheet1": df}
        except Exception:
            fh.seek(0)
            df = pd.read_csv(fh, dtype=str, keep_default_na=False, on_bad_lines='skip', encoding="latin-1")
            return {"Sheet1": df}

    # xlsx / xls / xlt — use xlrd engine for .xls/.xlt files
    # Some .xls/.xlt files are actually tab-separated text or HTML — handle all cases
    engine = "xlrd" if name.endswith((".xls", ".xlt")) else "openpyxl"
    try:
        xl = pd.ExcelFile(fh, engine=engine)
        return {sheet: xl.parse(sheet, dtype=str, keep_default_na=False) for sheet in xl.sheet_names}
    except Exception:
        # Fallback: try reading as tab-separated text
        fh.seek(0)
        raw = fh.read()
        # Detect HTML disguised as .xls — Proviso sometimes exports HTML
        if b"<!DOCTYPE" in raw[:500] or b"<html" in raw[:500]:
            try:
                tables = pd.read_html(io.BytesIO(raw))
                if tables:
                    df = tables[0].fillna("").astype(str)
                    # Check if it's a blank form (no real data)
                    if df.shape[0] < 5:
                        log.warning(f"    WARNING: {name} appears to be an empty HTML form — re-download required")
                        return {}
                    return {"Sheet1": df}
            except Exception:
                pass
            return {}
        fh.seek(0)
        try:
            df = pd.read_csv(fh, sep="\t", dtype=str, keep_default_na=False, encoding="utf-8", on_bad_lines='skip')
            return {"Sheet1": df}
        except Exception:
            fh.seek(0)
            df = pd.read_csv(fh, sep="\t", dtype=str, keep_default_na=False, encoding="latin-1", on_bad_lines='skip')
            return {"Sheet1": df}

def read_file_rows(drive, file_meta, financial_sheet=False):
    """
    Downloads file and returns a single DataFrame.
    If financial_sheet=True, picks the sheet whose name contains 'financial'.
    Otherwise picks the largest sheet (most cells).
    """
    sheets = download_file(drive, file_meta)
    if not sheets:
        return pd.DataFrame()

    if financial_sheet:
        for name, df in sheets.items():
            if "financial" in name.lower():
                return df
        # fallback: first sheet
        return next(iter(sheets.values()))

    # Pick largest sheet
    return max(sheets.values(), key=lambda df: df.shape[0] * df.shape[1])

def read_first_sheet(drive, file_meta):
    """Always returns the first sheet — used for file type detection."""
    sheets = download_file(drive, file_meta)
    if not sheets:
        return pd.DataFrame()
    return next(iter(sheets.values()))

# ─── GOOGLE SHEETS HELPERS ────────────────────────────────────────────────────

def sheets_read(sheets_svc, spreadsheet_id, range_):
    """Read a range from a Sheet. Returns list of lists."""
    resp = sheets_svc.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_
    ).execute()
    return resp.get("values", [])

def sheets_write(sheets_svc, spreadsheet_id, range_, values):
    """Write values to a Sheet range."""
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_,
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()

def sheets_append(sheets_svc, spreadsheet_id, range_, values):
    """Append rows to a Sheet."""
    sheets_svc.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_,
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": values},
    ).execute()

def get_sheet_as_df(sheets_svc, spreadsheet_id, tab_name):
    """Read a Sheet tab into a DataFrame."""
    data = sheets_read(sheets_svc, spreadsheet_id, tab_name)
    if not data or len(data) < 2:
        return pd.DataFrame()
    return pd.DataFrame(data[1:], columns=data[0])

def upsert_sheet_row(sheets_svc, spreadsheet_id, tab_name, match_cols, row_data):
    """
    Write a row to a Sheet tab. Updates existing row if match_cols match, else appends.
    match_cols: dict of {col_name: value} to find existing row.
    row_data: dict of all column values.
    """
    data = sheets_read(sheets_svc, spreadsheet_id, tab_name)
    if not data:
        return
    headers = data[0]
    
    # Find existing row
    row_idx = None
    for i, row in enumerate(data[1:], start=2):
        row_dict = dict(zip(headers, row))
        if all(str(row_dict.get(k, "")) == str(v) for k, v in match_cols.items()):
            row_idx = i
            break

    values = [[str(row_data.get(h, "")) for h in headers]]
    if row_idx:
        range_ = f"{tab_name}!A{row_idx}"
        sheets_write(sheets_svc, spreadsheet_id, range_, values)
    else:
        sheets_append(sheets_svc, spreadsheet_id, f"{tab_name}!A1", values)

# ─── DATE HELPERS ─────────────────────────────────────────────────────────────

def to_date(v) -> date | None:
    """Parse a value to a date. Returns None if unparseable."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, (date, datetime)):
        return v.date() if isinstance(v, datetime) else v
    s = str(v).strip()
    if not s or s in ("nan", "None", ""):
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    try:
        return pd.to_datetime(s, dayfirst=False).date()
    except Exception:
        return None

def last_day_of_month(month_str: str) -> date:
    """'2026-02' → date(2026, 2, 28)"""
    y, m = map(int, month_str.split("-"))
    return date(y, m, monthrange(y, m)[1])

def prior_month_str(month_str: str) -> str:
    """'2026-03' → '2026-02'"""
    y, m = map(int, month_str.split("-"))
    m -= 1
    if m == 0:
        m, y = 12, y - 1
    return f"{y}-{m:02d}"

def get_prior_month_str() -> str:
    today = date.today()
    m = today.month - 1 or 12
    y = today.year if today.month > 1 else today.year - 1
    return f"{y}-{m:02d}"

# ─── FILE TYPE DETECTION ─────────────────────────────────────────────────────

def identify_file_type(df: pd.DataFrame) -> str:
    """
    Identify file type from DataFrame content.
    Returns: 'pl' | 'status' | 'app' | 'apr' | 'disb' | 'unknown'
    """
    if df.empty:
        return "unknown"

    # P&L: column B (index 1) has 'exchange rate' AND 'interest income' within first 15 rows
    sample = df.iloc[:15]
    if df.shape[1] > 1:
        col_b = sample.iloc[:, 1].astype(str).str.lower()
        has_exchange = col_b.str.contains("exchange rate").any()
        has_interest = col_b.str.contains("interest income").any()
        if has_exchange and has_interest:
            return "pl"

    # Use header row for Proviso files
    headers = [str(v).strip().lower() for v in df.columns]
    h = "|".join(headers)

    if "overdue_age" in h or "overdue age" in h or "oamount" in h or "o_amount" in h:
        return "status"

    if ("principal" in h or "loan amount" in h) and any("date" in hh for hh in headers):
        return "disb"

    # APR check must come before APP — approval files also have a team column
    # Identify by presence of granted/requested amount columns or filename hint
    has_granted = "amt.granted" in h or "amt granted" in h or "amount granted" in h or "granted" in h
    has_requested = "amt.requested" in h or "amt requested" in h or "amount requested" in h or "requested" in h
    if has_granted or has_requested:
        return "apr"

    team_match = any(hh in ("teamid", "team_id", "team id", "team") for hh in headers)
    if team_match and any("date" in hh for hh in headers):
        return "app"

    if "approv" in h and "principal" not in h:
        return "apr"

    return "unknown"

def identify_files(drive, folder_id) -> dict:
    """
    Scan all files in a folder and return {type: file_meta}.
    Most-recently-modified file wins per type.
    """
    files = get_spreadsheet_files(drive, folder_id)
    identified = {"pl": None, "app": None, "apr": None, "disb": None, "status": None}

    for file_meta in files:
        log.info(f"    Reading: {file_meta['name']}")
        df = read_first_sheet(drive, file_meta)
        ftype = identify_file_type(df)
        if ftype == "unknown":
            log.info(f"    UNRECOGNISED: {file_meta['name']}")
        elif identified[ftype] is None:
            identified[ftype] = file_meta
            log.info(f"    {ftype.upper()}: {file_meta['name']}")
        else:
            log.info(f"    SKIPPED (duplicate {ftype}): {file_meta['name']}")

    return identified

# ─── P&L PARSER ──────────────────────────────────────────────────────────────

def parse_pl(drive, file_meta, month_str: str) -> dict:
    """
    Anchor-based P&L parser. Mirrors GAS parsePL() logic.
    Revenue = Interest Income + Processing Fees
    Opex    = Total Operating Costs + Interest Expense
    """
    df = read_file_rows(drive, file_meta, financial_sheet=True)
    if df.empty:
        raise ValueError("P&L file is empty")

    _, month = map(int, month_str.split("-"))
    rows = df.values.tolist()
    n_rows, n_cols = len(rows), max(len(r) for r in rows)

    def cell(r, c):
        try:
            v = rows[r][c]
            return v if v is not None else ""
        except IndexError:
            return ""

    def to_float(v):
        try:
            return float(str(v).replace(",", "")) if str(v).strip() not in ("", "nan") else 0.0
        except (ValueError, TypeError):
            return 0.0

    # Step 1: find USD section start (first column labelled 'USD')
    usd_start_col = -1
    for r in range(min(6, n_rows)):
        for c in range(n_cols):
            if str(cell(r, c)).strip() == "USD":
                usd_start_col = c
                break
        if usd_start_col >= 0:
            break
    if usd_start_col < 0:
        raise ValueError("USD section not found in P&L")

    # Step 2: find date header row (has parseable dates in USD section)
    header_row_idx = -1
    date_row = None
    for r in range(min(8, n_rows)):
        for c in range(usd_start_col, n_cols):
            d = to_date(cell(r, c))
            if d:
                header_row_idx = r
                date_row = rows[r]
                break
        if header_row_idx >= 0:
            break
    if header_row_idx < 0:
        raise ValueError("Date header row not found in P&L")

    # Step 3: find target month column (match by month only)
    usd_date_cols = [(c, to_date(date_row[c])) for c in range(usd_start_col, n_cols)
                     if c < len(date_row) and to_date(date_row[c])]

    col = -1
    for c, d in usd_date_cols:
        if d.month == month:
            col = c
            break
    if col < 0:
        raise ValueError(f"Month {month_str} column not found in P&L")

    # Prior = closest date column to the left in USD section
    prior_col = -1
    for c in range(col - 1, usd_start_col - 1, -1):
        if c < len(date_row) and to_date(date_row[c]):
            prior_col = c
            break

    # Step 4: exchange rate from LC section
    lc_date_cols = [(c, to_date(date_row[c])) for c in range(0, usd_start_col)
                    if c < len(date_row) and to_date(date_row[c])]

    lc_col = -1
    for c, d in lc_date_cols:
        if d.month == month:
            lc_col = c
            break
    if lc_col < 0 and lc_date_cols:
        usd_pos = next((i for i, (c, _) in enumerate(usd_date_cols) if c == col), -1)
        if 0 <= usd_pos < len(lc_date_cols):
            lc_col = lc_date_cols[usd_pos][0]

    exchange_rate = 1.0
    for r in range(min(10, n_rows)):
        if "exchange rate" in str(cell(r, 1)).lower():
            if lc_col >= 0:
                v = to_float(cell(r, lc_col))
                if v > 0:
                    exchange_rate = v
            break

    # Step 5: walk rows with anchors
    def cv(r): return to_float(cell(r, col))
    def pv(r): return to_float(cell(r, prior_col)) if prior_col >= 0 else 0.0

    result = {
        "exchangeRate": exchange_rate,
        "interestIncome": 0, "priorInterestIncome": 0,
        "interestExpense": 0, "priorInterestExpense": 0,
        "processingFees": 0, "priorProcessingFees": 0,
        "totalOpexLine":  0, "priorTotalOpexLine":  0,
        "operatingIncome":0, "priorOperatingIncome":0,
        "staffCost":      0, "priorStaffCost":      0,
        "fuelCost":       0, "priorFuelCost":       0,
        "vehicleCost":    0, "priorVehicleCost":    0,
        "provision":      0, "priorProvision":      0,
        "opexLines":      [],
    }

    BEFORE, REVENUE, OPEX = 0, 1, 2
    state, ie_found = BEFORE, False

    for r in range(header_row_idx + 1, n_rows):
        raw = str(cell(r, 1)).strip()
        lbl = raw.lower()
        if not raw:
            continue

        if "income statement" in lbl:
            state = REVENUE
            continue
        if state == BEFORE:
            continue

        # Stop anchor
        if "operating income" in lbl and "total" not in lbl:
            result["operatingIncome"]      = cv(r)
            result["priorOperatingIncome"] = pv(r)
            break

        # Provision anchor → switch to opex
        if "provision for loan" in lbl:
            c_v, p_v = cv(r), pv(r)
            result["provision"]      = -abs(c_v) if c_v != 0 else c_v
            result["priorProvision"] = -abs(p_v) if p_v != 0 else p_v
            state = OPEX
            continue

        if state == REVENUE:
            if "interest income" in lbl and "net" not in lbl and "expense" not in lbl:
                result["interestIncome"]      = cv(r)
                result["priorInterestIncome"] = pv(r)
            if "interest expense" in lbl and not ie_found:
                result["interestExpense"]      = cv(r)
                result["priorInterestExpense"] = pv(r)
                ie_found = True
            if "processing fees" in lbl:
                result["processingFees"]      = cv(r)
                result["priorProcessingFees"] = pv(r)

        if state == OPEX:
            if "total operating" in lbl:
                result["totalOpexLine"]      = cv(r)
                result["priorTotalOpexLine"] = pv(r)
                continue
            c_v, p_v = cv(r), pv(r)
            if c_v != 0 or p_v != 0:
                display_name = _normalize_opex_name(raw)
                existing = next((l for l in result["opexLines"] if l["name"] == display_name), None)
                if existing:
                    existing["current"] += c_v
                    existing["prior"]   += p_v
                else:
                    result["opexLines"].append({"name": display_name, "current": c_v, "prior": p_v})

                if any(key in lbl for key in CONFIG["PL_STAFF_KEYS"]):
                    result["staffCost"] += c_v; result["priorStaffCost"] += p_v
                if any(key in lbl for key in CONFIG["PL_FUEL_KEYS"]):
                    result["fuelCost"] += c_v; result["priorFuelCost"] += p_v
                if any(key in lbl for key in CONFIG["PL_VEHICLE_KEYS"]):
                    result["vehicleCost"] += c_v; result["priorVehicleCost"] += p_v

    # Derived totals
    result["revenue"]      = result["interestIncome"]      + result["processingFees"]
    result["priorRevenue"] = result["priorInterestIncome"] + result["priorProcessingFees"]
    result["opex"]         = result["totalOpexLine"]      + result["interestExpense"]
    result["priorOpex"]    = result["priorTotalOpexLine"] + result["priorInterestExpense"]

    return result

def _normalize_opex_name(raw: str) -> str:
    lbl = raw.lower()
    if "bank" in lbl or "momo" in lbl:
        if "disbursement" in lbl: return "Bank/MOMO Charges - Disbursement"
        if "repayment"    in lbl: return "Bank/MOMO Charges - Repayment"
        return "Bank/MOMO Charges - Other"
    return raw

# ─── PROVISO PARSERS ─────────────────────────────────────────────────────────

def _find_col(headers, *candidates, fallback=None):
    """Find the first header index matching any candidate string."""
    for h, i in [(h, i) for i, h in enumerate(headers)]:
        for cand in candidates:
            if cand in h:
                return i
    return fallback

def parse_app_file(drive, file_meta, month_str: str) -> dict | None:
    df = read_file_rows(drive, file_meta)
    if df.empty or len(df) < 2:
        return None
    year, month = map(int, month_str.split("-"))
    CM = CONFIG["COLUMN_MAPS"]
    headers = [str(v).strip().lower() for v in df.columns]

    date_idx = CM["app_date_index"]
    team_idx = CM["app_team_index"]
    area_idx = CM["app_area_index"]
    found_exact_team = False
    for i, h in enumerate(headers):
        if "date" in h and "expir" not in h and "birth" not in h:
            date_idx = i
        if h == "team":
            team_idx = i
            found_exact_team = True
        elif h in ("teamid", "team_id", "team id") and not found_exact_team:
            team_idx = i
        if h in ("area",) or "area_name" in h:
            area_idx = i

    log.info(f"    APP cols — date:{date_idx}({headers[date_idx] if date_idx < len(headers) else '?'}) "
             f"team:{team_idx} area:{area_idx}")

    total_apps = 0
    teams, areas = set(), set()

    for _, row in df.iterrows():
        vals = row.tolist()
        d = to_date(vals[date_idx] if date_idx < len(vals) else None)
        if not d or d.month != month or d.year != year:
            continue
        total_apps += 1
        team = str(vals[team_idx]).strip() if team_idx < len(vals) else ""
        area = str(vals[area_idx]).strip() if area_idx < len(vals) else ""
        if team and team.lower() not in ("nan", ""):
            teams.add(team)
        if area and area.lower() not in ("nan", ""):
            areas.add(area)

    num_teams    = len(teams) or 1
    apps_per_team = round(total_apps / num_teams) if num_teams else 0
    return {"totalApps": total_apps, "appsPerTeam": apps_per_team,
            "numTeams": num_teams, "numAreas": len(areas)}

def parse_apr_file(drive, file_meta, month_str: str) -> dict | None:
    df = read_file_rows(drive, file_meta)
    if df.empty or len(df) < 2:
        return None
    year, month = map(int, month_str.split("-"))
    CM = CONFIG["COLUMN_MAPS"]
    headers = [str(v).strip().lower() for v in df.columns]

    date_idx = CM["apr_date_index"]
    for i, h in enumerate(headers):
        if "date" in h and "expir" not in h:
            date_idx = i

    total_approvals = 0
    for _, row in df.iterrows():
        vals = row.tolist()
        d = to_date(vals[date_idx] if date_idx < len(vals) else None)
        if d and d.month == month and d.year == year:
            total_approvals += 1

    return {"totalApprovals": total_approvals}

def parse_disb_file(drive, file_meta, month_str: str, exchange_rate: float) -> dict | None:
    df = read_file_rows(drive, file_meta)
    if df.empty or len(df) < 2:
        return None
    year, month = map(int, month_str.split("-"))
    CM = CONFIG["COLUMN_MAPS"]
    headers = [str(v).strip().lower() for v in df.columns]
    fx = exchange_rate if exchange_rate > 1 else 1.0

    date_idx      = CM["disb_date_index"]
    principal_idx = CM["disb_principal_index"]
    team_idx      = CM["disb_team_index"]
    found_exact_team = False
    for i, h in enumerate(headers):
        if "date" in h and "expir" not in h:
            date_idx = i
        if h == "principal" or "loan amount" in h:
            principal_idx = i
        if h == "team":
            team_idx = i
            found_exact_team = True
        elif h in ("teamid", "team_id", "team id") and not found_exact_team:
            team_idx = i

    log.info(f"    DISB cols — date:{date_idx} principal:{principal_idx} team:{team_idx}")

    total_local, count = 0.0, 0
    teams = set()

    for _, row in df.iterrows():
        vals = row.tolist()
        d = to_date(vals[date_idx] if date_idx < len(vals) else None)
        if not d:
            continue
        principal = _to_float(vals[principal_idx] if principal_idx < len(vals) else 0)
        if principal <= 0:
            continue
        if d.month == month and d.year == year:
            total_local += principal
            count += 1
            team = str(vals[team_idx]).strip() if team_idx < len(vals) else ""
            if team and team.lower() not in ("nan", ""):
                teams.add(team)

    num_teams        = len(teams) or 1
    total_disb_usd   = total_local / fx
    disb_per_team    = round(count / num_teams) if num_teams else 0
    avg_loan_size_usd = total_disb_usd / count if count else 0.0
    return {"totalDisbUsd": total_disb_usd, "count": count, "numTeams": num_teams,
            "disbPerTeam": disb_per_team, "avgLoanSizeUsd": avg_loan_size_usd}

def parse_status_file(drive, file_meta, cc: str, date_str: str,
                      exchange_rate: float, par_window_principal: float,
                      chronic_window_pni: float, sheets_svc) -> dict | None:
    df = read_file_rows(drive, file_meta)
    if df.empty or len(df) < 2:
        return None
    CM = CONFIG["COLUMN_MAPS"]
    headers = [str(v).strip().lower() for v in df.columns]
    fx = exchange_rate if exchange_rate > 1 else 1.0
    report_date = datetime.strptime(date_str, "%Y-%m-%d").date()

    principal_idx      = CM["stat_principal_index"]   # OAmount (outstanding P+I)
    orig_principal_idx = -1                            # Original principal disbursed
    interest_idx       = CM["stat_interest_index"]
    overdue_age_idx    = CM["stat_overdue_age_index"]
    disb_date_idx      = -1
    chronic_idx        = -1

    for i, h in enumerate(headers):
        if h in ("oamount", "o_amount"):
            principal_idx = i
        if h == "principal_balance":
            orig_principal_idx = i
        if h == "interest_balance":
            interest_idx = i
        elif "interest" in h and "rate" not in h and "income" not in h and "paid" not in h:
            if interest_idx == CM["stat_interest_index"]:
                interest_idx = i
        if h in ("overdue_age", "dpd") or "overdue_age" in h or "overdue age" in h:
            overdue_age_idx = i
        if h in ("date_granted", "disbursement_date", "disburse_date", "loan_date", "date_of_grant"):
            disb_date_idx = i
        if disb_date_idx < 0 and ("disb" in h or "grant" in h) and "date" in h:
            disb_date_idx = i
        if h in ("chronic_status", "chronic") or "chronic_status" in h or "loanstatus" in h:
            chronic_idx = i

    log.info(f"    STATUS cols — oamount:{principal_idx} origPrincipal:{orig_principal_idx} "
             f"overdueAge:{overdue_age_idx} disbDate:{disb_date_idx} chronic:{chronic_idx}")

    loan_book_local  = 0.0
    par30_num        = 0.0
    par30_fallback_den = 0.0
    current_chronic_pni = 0.0

    for _, row in df.iterrows():
        vals = row.tolist()
        oamount      = _to_float(vals[principal_idx]      if principal_idx      < len(vals) else 0)
        prin_bal     = _to_float(vals[orig_principal_idx] if orig_principal_idx >= 0 and orig_principal_idx < len(vals) else 0)
        overdue_age  = _to_float(vals[overdue_age_idx]    if overdue_age_idx    < len(vals) else 0)
        if oamount <= 0:
            continue

        # Loan book: OAmount filtered by overdue ≤ 60 days
        if overdue_age <= 60:
            loan_book_local += oamount

        # PAR 30 — uses Principal_Balance for both numerator and denominator
        disb_date = to_date(vals[disb_date_idx] if disb_date_idx >= 0 and disb_date_idx < len(vals) else None)
        if disb_date:
            days_since = (report_date - disb_date).days
            if 60 <= days_since <= 133:
                par30_fallback_den += prin_bal
                is_chronic = chronic_idx >= 0 and str(vals[chronic_idx]).strip().lower() == "chronic"
                if overdue_age > 30 and not is_chronic:
                    par30_num += prin_bal

        # Chronic numerator — uses OAmount (P+I outstanding)
        is_chronic = chronic_idx >= 0 and str(vals[chronic_idx] if chronic_idx < len(vals) else "").strip().lower() == "chronic"
        if is_chronic:
            current_chronic_pni += oamount

    # PAR 30 rate
    par30_den  = par_window_principal if par_window_principal > 0 else par30_fallback_den
    par30_rate = par30_num / par30_den if par30_den > 0 else 0.0

    # Chronic rate — subtract Jan 1 chronic baseline
    jan1_chronic_pni = _read_baseline_chronic(sheets_svc, cc)
    chronic_numerator = max(0, current_chronic_pni - jan1_chronic_pni)
    chronic_rate = chronic_numerator / chronic_window_pni if chronic_window_pni > 0 else 0.0
    log.info(f"    Chronic debug — current={round(current_chronic_pni):,} jan1={round(jan1_chronic_pni):,} "
             f"numerator={round(chronic_numerator):,} window_den={round(chronic_window_pni):,}")

    log.info(f"    Status: loanBook={fmt_usd(loan_book_local/fx)}, "
             f"PAR30={fmt_pct(par30_rate)}, Chronic={fmt_pct(chronic_rate)}")

    return {"loanBookUsd": loan_book_local / fx, "par30Rate": par30_rate, "chronicRate": chronic_rate}

def _read_baseline_chronic(sheets_svc, cc: str) -> float:
    """Read Jan 1 chronic OAmount from cached baseline sheet tab."""
    data = sheets_read(sheets_svc, CONFIG["SHEET_ID"], f"{cc}_Baseline")
    if not data or len(data) < 2:
        return 0.0
    total = 0.0
    for row in data[1:]:
        if len(row) < 2:
            continue
        oamt   = _to_float(row[1])
        status = str(row[3]).strip().lower() if len(row) > 3 else ""
        if status == "chronic":
            total += oamt
    return total

# ─── DISBURSEMENT WINDOW HELPER ───────────────────────────────────────────────

def get_disb_window_data(drive, sheets_svc, cc: str, month_str: str,
                         window_start: date, window_end: date) -> dict:
    """
    Sum disbursements within [window_start, window_end] from:
    1. The cached Disb2025 sheet tab
    2. All 2026+ monthly Drive folders up to month_str
    Returns {"principal": float, "pAndI": float}
    """
    total_principal, total_pni = 0.0, 0.0

    # 1. Read from cache sheet
    data = sheets_read(sheets_svc, CONFIG["SHEET_ID"], f"{cc}_Disb2025")
    if data and len(data) > 1:
        for row in data[1:]:
            if len(row) < 2:
                continue
            d = to_date(row[0])
            if not d or not (window_start <= d <= window_end):
                continue
            p   = _to_float(row[1])
            int_= _to_float(row[2]) if len(row) > 2 else 0.0
            total_principal += p
            total_pni       += p + int_

    # 2. Scan monthly Drive folders
    year, month = map(int, month_str.split("-"))
    root_files = list_folder(drive, CONFIG["DRIVE_ROOT"])
    month_folders = [
        f for f in root_files
        if f["mimeType"] == "application/vnd.google-apps.folder"
        and re.match(r"^\d{4}-\d{2}$", f["name"])
    ]

    for mf in month_folders:
        fy, fm = map(int, mf["name"].split("-"))
        if fy > year or (fy == year and fm > month):
            continue
        folder_month_end = date(fy, fm, monthrange(fy, fm)[1])
        if folder_month_end < window_start:
            continue

        cc_folder_id = find_subfolder(drive, mf["id"], cc)
        if not cc_folder_id:
            continue

        for file_meta in get_spreadsheet_files(drive, cc_folder_id):
            df_first = read_first_sheet(drive, file_meta)
            if identify_file_type(df_first) != "disb":
                continue
            df = read_file_rows(drive, file_meta)
            if df.empty:
                continue
            headers = [str(v).strip().lower() for v in df.columns]
            date_idx, prin_idx, int_idx = 0, 4, -1
            for i, h in enumerate(headers):
                if h in ("input_date", "date_granted", "disbursement_date", "loan_date"):
                    date_idx = i
                elif "date" in h and "expir" not in h and "deduct" not in h and date_idx == 0:
                    date_idx = i
                if h == "principal":
                    prin_idx = i
                if h in ("interest", "interest_amount"):
                    int_idx = i
            for _, row in df.iterrows():
                vals = row.tolist()
                d = to_date(vals[date_idx] if date_idx < len(vals) else None)
                if not d or not (window_start <= d <= window_end):
                    continue
                p   = _to_float(vals[prin_idx] if prin_idx < len(vals) else 0)
                int_= _to_float(vals[int_idx]  if int_idx  >= 0 and int_idx < len(vals) else 0)
                total_principal += p
                total_pni       += p + int_

    return {"principal": total_principal, "pAndI": total_pni}

# ─── MAIN PIPELINE ───────────────────────────────────────────────────────────

def process_country_month(drive, sheets_svc, cc: str, month_str: str, folder_id: str) -> dict | None:
    """Process all files for one country in one month. Returns metrics dict or None."""
    log.info(f"\n── {cc} {month_str} ──")

    identified = identify_files(drive, folder_id)
    baseline_avail = _check_baseline(sheets_svc, cc)
    log.info(f"  BASELINE: {'found' if baseline_avail else 'not found'}")

    if not identified["pl"]:
        log.info("  SKIP: no P&L file")
        return None

    date_str    = last_day_of_month(month_str).strftime("%Y-%m-%d")
    report_date = last_day_of_month(month_str)
    year, month = map(int, month_str.split("-"))

    # P&L
    pl = parse_pl(drive, identified["pl"], month_str)
    log.info(f"  P&L: FX={pl['exchangeRate']}, Revenue={fmt_usd(pl['revenue'])}, "
             f"OI={fmt_usd(pl['operatingIncome'])}")

    # App / Apr / Disb
    app  = parse_app_file( drive, identified["app"],  month_str)           if identified["app"]  else None
    apr  = parse_apr_file( drive, identified["apr"],  month_str)           if identified["apr"]  else None
    disb = parse_disb_file(drive, identified["disb"], month_str, pl["exchangeRate"]) if identified["disb"] else None

    # PAR window: disbursed 60–133 days before report date
    par_window_start = report_date - timedelta(days=133)
    par_window_end   = report_date - timedelta(days=60)
    log.info(f"  PAR window: {par_window_start} → {par_window_end}")
    par_data = get_disb_window_data(drive, sheets_svc, cc, month_str, par_window_start, par_window_end)
    par_window_principal = par_data["principal"]
    log.info(f"  PAR window principal: {round(par_window_principal):,}")

    # Chronic window: disbursed 134 days before Jan 1 → 134 days before report date
    jan1 = date(year, 1, 1)
    chronic_window_start = jan1       - timedelta(days=134)
    chronic_window_end   = report_date - timedelta(days=134)
    log.info(f"  Chronic window: {chronic_window_start} → {chronic_window_end}")
    chr_data = get_disb_window_data(drive, sheets_svc, cc, month_str, chronic_window_start, chronic_window_end)
    chronic_window_pni = chr_data["pAndI"]
    log.info(f"  Chronic window P+I: {round(chronic_window_pni):,}")

    # Status
    status = None
    if identified["status"]:
        status = parse_status_file(
            drive, identified["status"], cc, date_str,
            pl["exchangeRate"], par_window_principal, chronic_window_pni, sheets_svc
        )

    return {
        "pl": pl, "app": app, "apr": apr, "disb": disb, "status": status
    }

def _check_baseline(sheets_svc, cc: str) -> bool:
    """Returns True if baseline tab has data rows."""
    data = sheets_read(sheets_svc, CONFIG["SHEET_ID"], f"{cc}_Baseline")
    return bool(data and len(data) > 1)

def process_month(month_str: str):
    """Main entry point — process all countries for a month."""
    log.info(f"\n{'='*60}")
    log.info(f"OYA Monthly — processing {month_str}")
    log.info(f"{'='*60}")

    drive, sheets_svc = get_services()
    ss_id = CONFIG["SHEET_ID"]

    # Find month folder in Drive
    root_items = list_folder(drive, CONFIG["DRIVE_ROOT"])
    month_folder = next(
        (f for f in root_items
         if f["mimeType"] == "application/vnd.google-apps.folder"
         and f["name"] == month_str),
        None
    )
    if not month_folder:
        raise ValueError(f"Month folder '{month_str}' not found in Drive root")

    results = {}

    for cc in CONFIG["COUNTRIES"]:
        cc_folder_id = find_subfolder(drive, month_folder["id"], cc)
        if not cc_folder_id:
            log.info(f"\n── {cc}: no folder found ──")
            continue
        files_in_folder = get_spreadsheet_files(drive, cc_folder_id)
        if not files_in_folder:
            log.info(f"\n── {cc}: folder empty ──")
            continue

        try:
            data = process_country_month(drive, sheets_svc, cc, month_str, cc_folder_id)
            if data:
                results[cc] = data
                store_results(sheets_svc, ss_id, cc, month_str, data)
                log.info(f"  ✓ {cc} stored to Sheet")
        except Exception as e:
            log.error(f"  ✗ {cc} ERROR: {e}", exc_info=True)

    # Build and commit JSON
    payload = build_json(sheets_svc, ss_id, month_str)
    if payload:
        _commit_to_github(month_str, payload)

    log.info(f"\n{'='*60}")
    log.info(f"Done — {month_str}")
    log.info(f"{'='*60}\n")

def _get_prior_row(sheets_svc, ss_id: str, cc: str, month_str: str) -> dict:
    """Fetch the stored row for the prior month for a country."""
    prev = prior_month_str(month_str)
    df = get_sheet_as_df(sheets_svc, ss_id, f"{cc}_Monthly")
    if df.empty:
        return {}
    row = df[df["month"].astype(str) == prev]
    if row.empty:
        return {}
    return {k: _to_float(v) for k, v in row.iloc[0].to_dict().items()}

# ─── RESULT STORAGE ──────────────────────────────────────────────────────────

def store_results(sheets_svc, ss_id: str, cc: str, month_str: str, data: dict):
    """Write processed results to the country's Monthly sheet tab."""
    pl     = data["pl"]
    app    = data.get("app")   or {}
    apr    = data.get("apr")   or {}
    disb   = data.get("disb")  or {}
    status = data.get("status") or {}

    approval_rate = (apr.get("totalApprovals", 0) / app["totalApps"]
                     if app.get("totalApps") else 0)
    num_teams = app.get("numTeams") or disb.get("numTeams") or 0

    # Load prior month row for operational prior fields
    pr = _get_prior_row(sheets_svc, ss_id, cc, month_str)

    row = {
        "month":                  month_str,
        "exchange_rate":          pl["exchangeRate"],
        "total_apps":             app.get("totalApps", 0),
        "apps_per_team":          app.get("appsPerTeam", 0),
        "num_teams":              num_teams,
        "num_areas":              app.get("numAreas", 0),
        "prior_apps_per_team":    pr.get("apps_per_team", 0),
        "prior_total_apps":       pr.get("total_apps", 0),
        "prior_num_teams":        pr.get("num_teams", 0),
        "total_approvals":        apr.get("totalApprovals", 0),
        "prior_total_approvals":  pr.get("total_approvals", 0),
        "approval_rate":          approval_rate,
        "prior_approval_rate":    pr.get("approval_rate", 0),
        "total_disb_usd":         disb.get("totalDisbUsd", 0),
        "disb_per_team":          disb.get("disbPerTeam", 0),
        "avg_loan_size_usd":      disb.get("avgLoanSizeUsd", 0),
        "num_loans":              disb.get("count", 0),
        "prior_total_disb_usd":   pr.get("total_disb_usd", 0),
        "prior_disb_per_team":    pr.get("disb_per_team", 0),
        "prior_avg_loan_size_usd":pr.get("avg_loan_size_usd", 0),
        "loan_book_usd":          status.get("loanBookUsd", 0),
        "prior_loan_book_usd":    pr.get("loan_book_usd", 0),
        "par30_rate":             status.get("par30Rate", 0),
        "prior_par30_rate":       pr.get("par30_rate", 0),
        "chronic_rate":           status.get("chronicRate", 0),
        "prior_chronic_rate":     pr.get("chronic_rate", 0),
        "revenue_usd":            pl["revenue"],
        "prior_revenue_usd":      pl["priorRevenue"],
        "provision_usd":          pl["provision"],
        "prior_provision_usd":    pl["priorProvision"],
        "opex_usd":               pl["opex"],
        "prior_opex_usd":         pl["priorOpex"],
        "operating_income_usd":   pl["operatingIncome"],
        "prior_operating_income_usd": pl["priorOperatingIncome"],
        "interest_income_usd":    pl["interestIncome"],
        "prior_interest_income_usd": pl["priorInterestIncome"],
        "processing_fees_usd":    pl["processingFees"],
        "prior_processing_fees_usd": pl["priorProcessingFees"],
        "staff_cost_usd":         pl["staffCost"],
        "prior_staff_cost_usd":   pl["priorStaffCost"],
        "fuel_cost_usd":          pl["fuelCost"],
        "prior_fuel_cost_usd":    pl["priorFuelCost"],
        "vehicle_cost_usd":       pl["vehicleCost"],
        "prior_vehicle_cost_usd": pl["priorVehicleCost"],
    }

    tab = f"{cc}_Monthly"
    upsert_sheet_row(sheets_svc, ss_id, tab, {"month": month_str}, row)

# ─── JSON BUILDER ─────────────────────────────────────────────────────────────

def build_json(sheets_svc, ss_id: str, month_str: str) -> dict | None:
    """Read all country results for month_str from Sheet and return JSON payload."""
    payload = {"monthStr": month_str}
    any_data = False

    for cc in CONFIG["COUNTRIES"]:
        df = get_sheet_as_df(sheets_svc, ss_id, f"{cc}_Monthly")
        if df.empty:
            continue
        row = df[df["month"].astype(str) == month_str]
        if row.empty:
            continue
        any_data = True
        r = row.iloc[0].to_dict()
        payload[cc] = {k: _safe_num(v) for k, v in r.items() if k != "month"}

    payload["commentary"] = {"apps": None, "efficiency": None,
                             "disbursements": None, "loan_book": None}
    return payload if any_data else None

def _safe_num(v):
    """Convert sheet values to clean Python numbers or strings."""
    try:
        f = float(v)
        return int(f) if f == int(f) else round(f, 6)
    except (ValueError, TypeError):
        return str(v) if v not in (None, "") else None

# ─── GITHUB ───────────────────────────────────────────────────────────────────

def _commit_to_github(month_str: str, payload: dict):
    """Commit data/{month_str}.json to GitHub."""
    token = CONFIG["GITHUB"]["TOKEN"]
    if not token:
        log.warning("GITHUB_TOKEN not set — skipping GitHub commit")
        return

    owner  = CONFIG["GITHUB"]["OWNER"]
    repo   = CONFIG["GITHUB"]["REPO"]
    branch = CONFIG["GITHUB"]["BRANCH"]
    path   = f"data/{month_str}.json"
    url    = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"

    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
    }

    # Get existing SHA if file exists
    sha = None
    resp = requests.get(url, headers=headers)
    if resp.status_code == 200:
        sha = resp.json().get("sha")

    content = base64.b64encode(json.dumps(payload, indent=2).encode()).decode()
    body = {"message": f"Monthly data: {month_str}", "content": content, "branch": branch}
    if sha:
        body["sha"] = sha

    put_resp = requests.put(url, headers=headers, json=body)
    if put_resp.status_code in (200, 201):
        log.info(f"  ✓ Committed data/{month_str}.json to GitHub")
    else:
        log.error(f"  ✗ GitHub commit failed ({put_resp.status_code}): {put_resp.text[:300]}")

# ─── BASELINE SETUP ───────────────────────────────────────────────────────────

def setup_baseline_data():
    """
    One-time: cache Jan 1 STATUS and 2025 DISBURSEMENTS files to Sheet tabs.
    Run this once after uploading baseline files to Drive.
    """
    log.info("Setting up baseline data cache...")
    drive, sheets_svc = get_services()

    for cc in CONFIG["COUNTRIES"]:
        log.info(f"\n── {cc} ──")
        _cache_baseline_status(drive, sheets_svc, cc)
        _cache_baseline_disb(drive, sheets_svc, cc)

    log.info("\n✓ All baseline data cached.")

def _cache_baseline_status(drive, sheets_svc, cc: str):
    baseline_id = _get_folder_id_from_sheet(sheets_svc, f"BASELINE_{cc}")
    if not baseline_id:
        log.info(f"  {cc}: BASELINE_{cc} folder ID not found in Monthly_Config tab")
        return
    files = get_spreadsheet_files(drive, baseline_id)
    if not files:
        log.info(f"  {cc}: No Jan 1 STATUS file in baseline folder")
        return

    file_meta = files[0]
    log.info(f"  {cc}: Caching Jan 1 STATUS: {file_meta['name']}")
    df = read_file_rows(drive, file_meta)
    if df.empty:
        return

    headers = [str(v).strip().lower() for v in df.columns]
    prin_idx, int_idx, date_idx, chronic_idx = -1, -1, -1, -1
    for i, h in enumerate(headers):
        if h in ("oamount", "o_amount"):
            prin_idx = i
        elif h == "principal_balance" and prin_idx < 0:
            prin_idx = i
        if h == "interest_balance":
            int_idx = i
        elif "interest" in h and "rate" not in h and "income" not in h and int_idx < 0:
            int_idx = i
        if h in ("date_granted", "disbursement_date", "loan_date"):
            date_idx = i
        if date_idx < 0 and ("disb" in h or "grant" in h) and "date" in h:
            date_idx = i
        if h in ("chronic_status", "chronic") or "chronic_status" in h:
            chronic_idx = i

    rows = []
    for _, row in df.iterrows():
        vals = row.tolist()
        p    = _to_float(vals[prin_idx] if prin_idx >= 0 and prin_idx < len(vals) else 0)
        int_ = _to_float(vals[int_idx]  if int_idx  >= 0 and int_idx  < len(vals) else 0)
        # If we found OAmount use it directly, otherwise sum P+I
        oamt = p if headers[prin_idx] in ("oamount", "o_amount") else p + int_
        if oamt <= 0:
            continue
        d       = to_date(vals[date_idx] if date_idx >= 0 and date_idx < len(vals) else None)
        chronic = str(vals[chronic_idx] if chronic_idx >= 0 and chronic_idx < len(vals) else "").strip().lower()
        rows.append([d.strftime("%Y-%m-%d") if d else "", oamt, 0, chronic])

    tab = f"{cc}_Baseline"
    sheets_svc.spreadsheets().values().clear(
        spreadsheetId=CONFIG["SHEET_ID"], range=f"{tab}!A2:D"
    ).execute()
    if rows:
        sheets_write(sheets_svc, CONFIG["SHEET_ID"], f"{tab}!A2", rows)
    log.info(f"  {cc}_Baseline: {len(rows)} loans cached")

def _cache_baseline_disb(drive, sheets_svc, cc: str):
    disb_id = _get_folder_id_from_sheet(sheets_svc, f"DISB2025_{cc}")
    if not disb_id:
        log.info(f"  {cc}: DISB2025_{cc} folder ID not found in Monthly_Config tab")
        return
    files = get_spreadsheet_files(drive, disb_id)
    if not files:
        log.info(f"  {cc}: No 2025 DISB file found")
        return

    file_meta = files[0]
    log.info(f"  {cc}: Caching 2025 DISB: {file_meta['name']}")
    df = read_file_rows(drive, file_meta)
    if df.empty:
        return

    headers = [str(v).strip().lower() for v in df.columns]
    date_idx, prin_idx, int_idx = -1, -1, -1
    for i, h in enumerate(headers):
        if h in ("input_date", "date_granted", "disbursement_date", "loan_date"):
            date_idx = i
        elif "date" in h and "expir" not in h and "deduct" not in h and date_idx < 0:
            date_idx = i
        if h == "principal":
            prin_idx = i
        if h in ("interest", "interest_amount"):
            int_idx = i

    rows = []
    for _, row in df.iterrows():
        vals = row.tolist()
        d = to_date(vals[date_idx] if date_idx >= 0 and date_idx < len(vals) else None)
        if not d:
            continue
        p   = _to_float(vals[prin_idx] if prin_idx >= 0 and prin_idx < len(vals) else 0)
        int_= _to_float(vals[int_idx]  if int_idx  >= 0 and int_idx  < len(vals) else 0)
        if p <= 0:
            continue
        rows.append([d.strftime("%Y-%m-%d"), p, int_])

    tab = f"{cc}_Disb2025"
    sheets_svc.spreadsheets().values().clear(
        spreadsheetId=CONFIG["SHEET_ID"], range=f"{tab}!A2:C"
    ).execute()
    if rows:
        sheets_write(sheets_svc, CONFIG["SHEET_ID"], f"{tab}!A2", rows)
    log.info(f"  {cc}_Disb2025: {len(rows)} disbursements cached")

def _get_folder_id_from_sheet(sheets_svc, key: str) -> str | None:
    """Look up a folder ID from the Monthly_Config sheet tab."""
    data = sheets_read(sheets_svc, CONFIG["SHEET_ID"], "Monthly_Config")
    if not data or len(data) < 2:
        return None
    for row in data[1:]:
        if len(row) >= 2 and str(row[0]) == key:
            return str(row[1])
    return None

# ─── UTILITY ─────────────────────────────────────────────────────────────────

def _to_float(v) -> float:
    try:
        return float(str(v).replace(",", "")) if str(v).strip() not in ("", "nan", "None") else 0.0
    except (ValueError, TypeError):
        return 0.0

def fmt_usd(n: float) -> str:
    if n is None: return "—"
    n = float(n); sign = "-" if n < 0 else ""; a = abs(n)
    if a >= 1e6: return f"{sign}${a/1e6:.2f}M"
    if a >= 1e3: return f"{sign}${round(a/1e3)}K"
    return f"{sign}${round(a):,}"

def fmt_pct(n: float) -> str:
    return f"{float(n)*100:.1f}%" if n is not None else "—"

# ─── ENTRY POINT ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if arg == "setup-baseline":
            setup_baseline_data()
        elif re.match(r"^\d{4}-\d{2}$", arg):
            process_month(arg)
        else:
            print(f"Usage: python process.py [YYYY-MM | setup-baseline]")
            sys.exit(1)
    else:
        # Default: process prior month
        month = get_prior_month_str()
        log.info(f"No month specified — defaulting to prior month: {month}")
        process_month(month)
