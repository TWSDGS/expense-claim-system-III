import os
from typing import Dict, List
import pandas as pd
from openpyxl import Workbook

# Canonical column order
COLUMNS: List[str] = [
    "id",
    "status",
    "filler_name",
    "form_date",
    "plan_code",
    "purpose_desc",
    "payment_mode",
    "payee_type",
    "employee_name",
    "employee_no",
    "vendor_name",
    "vendor_address",
    "vendor_payee_name",
    "is_advance_offset",
    "advance_amount",
    "offset_amount",
    "balance_refund_amount",
    "supplement_amount",
    "receipt_no",
    "amount_untaxed",
    "tax_amount",
    "amount_total",
    "handler_name",
    "project_manager_name",
    "dept_manager_name",
    "accountant_name",
    "attachments",
    "created_at",
    "updated_at",
    "submitted_at",
]

SHEET_NAME = "vouchers"

def ensure_workbook(xlsx_path: str):
    os.makedirs(os.path.dirname(xlsx_path), exist_ok=True)
    if os.path.exists(xlsx_path):
        return
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(COLUMNS)
    wb.save(xlsx_path)

def _read_df(xlsx_path: str) -> pd.DataFrame:
    if not os.path.exists(xlsx_path):
        ensure_workbook(xlsx_path)
    try:
        df = pd.read_excel(xlsx_path, sheet_name=SHEET_NAME, dtype=str)
    except Exception:
        ensure_workbook(xlsx_path)
        df = pd.read_excel(xlsx_path, sheet_name=SHEET_NAME, dtype=str)

    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ""
    df = df[COLUMNS]
    return df.fillna("")

def load_all(xlsx_path: str) -> pd.DataFrame:
    df = _read_df(xlsx_path)
    if not df.empty and "id" in df.columns:
        df = df.sort_values("id", ascending=False, kind="mergesort").reset_index(drop=True)
    return df

def upsert_record(xlsx_path: str, record: Dict):
    df = _read_df(xlsx_path)
    rid = str(record.get("id", "")).strip()
    if not rid:
        raise ValueError("record.id is required")

    row = {c: str(record.get(c, "")) for c in COLUMNS}
    hit = df["id"] == rid
    if hit.any():
        df.loc[hit, COLUMNS] = pd.DataFrame([row])[COLUMNS].values
    else:
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)

    df = df.sort_values("id", ascending=False, kind="mergesort").reset_index(drop=True)

    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

def delete_record(xlsx_path: str, record_id: str):
    df = _read_df(xlsx_path)
    df = df[df["id"] != str(record_id)]
    df = df.sort_values("id", ascending=False, kind="mergesort").reset_index(drop=True)

    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)


# Compatibility wrapper: app.py imports `upsert` (older API name)
def upsert(xlsx_path: str, record: Dict):
    return upsert_record(xlsx_path, record)
