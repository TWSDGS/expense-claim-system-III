from __future__ import annotations

import os
import json
from typing import Dict, List, Optional

import pandas as pd

# We use gspread + google-auth to write to Google Sheets.
# Recommended auth: Service Account JSON file.
#
# Setup summary:
# 1) Create a Google Sheet (owned by twsdgs@gmail.com)
# 2) Create a Google Cloud project -> enable Google Sheets API
# 3) Create a Service Account -> download JSON key file
# 4) Share the Google Sheet with the service account email (Editor)
#
# Env vars supported:
#   GOOGLE_SHEET_ID             (required for cloud mode)
#   GOOGLE_WORKSHEET            (default: vouchers)
#   GOOGLE_SERVICE_ACCOUNT_FILE (path to service account json)
#   GOOGLE_SERVICE_ACCOUNT_JSON (raw json string; optional)
#
# Note: Google Sheets API operates on Google Sheets spreadsheets (not native .xlsx).
# You can still "File -> Download -> Microsoft Excel (.xlsx)" from the sheet at any time.

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def _get_gspread_client(service_account_file: str = "", service_account_json: str = ""):
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except Exception as e:
        raise RuntimeError("Missing dependency. Please install: gspread google-auth") from e

    if service_account_json:
        info = json.loads(service_account_json)
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)

    if not service_account_file:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_FILE is required (or GOOGLE_SERVICE_ACCOUNT_JSON).")

    if not os.path.exists(service_account_file):
        raise RuntimeError(f"Service account json not found: {service_account_file}")

    creds = Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
    return gspread.authorize(creds)

def ensure_worksheet(
    sheet_id: str,
    worksheet_name: str,
    columns: List[str],
    service_account_file: str = "",
    service_account_json: str = "",
) -> None:
    """Ensure worksheet exists and has header row."""
    gc = _get_gspread_client(service_account_file, service_account_json)
    sh = gc.open_by_key(sheet_id)

    try:
        ws = sh.worksheet(worksheet_name)
    except Exception:
        ws = sh.add_worksheet(title=worksheet_name, rows=2000, cols=max(10, len(columns)))

    values = ws.get_all_values()
    if not values:
        ws.append_row(columns, value_input_option="RAW")

def load_all_google(
    sheet_id: str,
    worksheet_name: str = "vouchers",
    columns: Optional[List[str]] = None,
    service_account_file: str = "",
    service_account_json: str = "",
) -> pd.DataFrame:
    gc = _get_gspread_client(service_account_file, service_account_json)
    sh = gc.open_by_key(sheet_id)
    ws = sh.worksheet(worksheet_name)

    values = ws.get_all_values()
    if not values:
        return pd.DataFrame(columns=columns or [])
    header = values[0]
    rows = values[1:]

    df = pd.DataFrame(rows, columns=header)
    if columns:
        for c in columns:
            if c not in df.columns:
                df[c] = ""
        df = df[columns]
    return df

def _find_row_index_by_id(ws, record_id: str, id_col_name: str = "id") -> Optional[int]:
    """Return 1-based row index in sheet (including header row)."""
    values = ws.get_all_values()
    if not values:
        return None
    header = values[0]
    if id_col_name not in header:
        return None
    id_idx = header.index(id_col_name)
    for i, row in enumerate(values[1:], start=2):
        if len(row) > id_idx and str(row[id_idx]).strip() == str(record_id).strip():
            return i
    return None

def upsert_record_google(
    sheet_id: str,
    payload: Dict,
    worksheet_name: str = "vouchers",
    columns: Optional[List[str]] = None,
    service_account_file: str = "",
    service_account_json: str = "",
) -> None:
    gc = _get_gspread_client(service_account_file, service_account_json)
    sh = gc.open_by_key(sheet_id)
    ws = sh.worksheet(worksheet_name)

    values = ws.get_all_values()
    if not values:
        if not columns:
            raise RuntimeError("Worksheet is empty and columns not provided.")
        ws.append_row(columns, value_input_option="RAW")
        header = columns
    else:
        header = values[0]

    if columns:
        missing = [c for c in columns if c not in header]
        if missing:
            header = header + missing
            # Rewrite the whole header row to avoid A1 column letter edge cases (>Z)
            ws.update("A1", [header], value_input_option="RAW")

    record_id = str(payload.get("id", "")).strip()
    if not record_id:
        raise RuntimeError("payload.id is required")

    row = [str(payload.get(col, "")) for col in header]
    row_idx = _find_row_index_by_id(ws, record_id, "id")
    if row_idx is None:
        ws.append_row(row, value_input_option="RAW")
    else:
        ws.update(f"A{row_idx}", [row], value_input_option="RAW")

def delete_record_google(
    sheet_id: str,
    record_id: str,
    worksheet_name: str = "vouchers",
    service_account_file: str = "",
    service_account_json: str = "",
) -> None:
    gc = _get_gspread_client(service_account_file, service_account_json)
    sh = gc.open_by_key(sheet_id)
    ws = sh.worksheet(worksheet_name)

    row_idx = _find_row_index_by_id(ws, str(record_id), "id")
    if row_idx is None:
        return
    ws.delete_rows(row_idx)

def build_sheet_url(sheet_id: str) -> str:
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}"
