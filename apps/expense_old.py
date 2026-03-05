# app.py
from __future__ import annotations

import json
import os
import re
import shutil
from datetime import datetime, date
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

from storage_excel import load_all as load_local_all, upsert as upsert_local, delete_record as delete_local_record
from storage_apps_script import AppsScriptError, list_records as cloud_list, upsert_record as cloud_upsert, delete_record as cloud_delete
from pdf_gen import build_pdf_bytes


APP_DIR = Path(__file__).resolve().parent
BASE_DIR = APP_DIR.parent
DATA_DIR = APP_DIR / "data"
ATTACH_DIR = DATA_DIR / "attachments"
CONFIG_PATH = DATA_DIR / "config.json"
LOCAL_XLSX = DATA_DIR / "vouchers.xlsx"
BG_IMAGE = BASE_DIR / "templates" / "voucher_bg.png"

DATA_DIR.mkdir(exist_ok=True, parents=True)
ATTACH_DIR.mkdir(exist_ok=True, parents=True)


# ----------------------------
# Helpers
# ----------------------------
def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")




def short_text(s: object, n: int = 10) -> str:
    """Return first n chars, append … if truncated."""
    try:
        t = "" if s is None else str(s)
    except Exception:
        t = ""
    t = t.strip()
    if n <= 0:
        return ""
    return t if len(t) <= n else (t[:n] + "…")

def parse_sheet_id(s: str) -> str:
    """Accept spreadsheet ID or a Google Sheets URL and return the ID."""
    s = (s or "").strip()
    if not s:
        return ""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", s)
    if m:
        return m.group(1)
    # If user pasted a gid URL or bare id, keep as-is (simple validation)
    return s


def normalize_apps_script_url(s: str) -> str:
    """Accept full Web App URL or Deployment ID and return Web App URL."""
    s = (s or "").strip()
    if not s:
        return ""
    if s.startswith("http://") or s.startswith("https://"):
        return s
    # looks like deployment id
    if s.startswith("AKfy"):
        return f"https://script.google.com/macros/s/{s}/exec"
    return s


def _read_json(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


DEFAULT_CONFIG = {
    "backend": "google",  # local | google
    "google": {
        "spreadsheet_id": "1i8Iw8dTfrKGpCOdxMXl5d2QMgOD7VbA84UEPRjBc_zw",
        "submit_sheet_name": "申請表單",
        "draft_sheet_name": "草稿列表",
        "apps_script_url": "https://script.google.com/macros/s/AKfycbxJjuJPg6CXECoeKTm4o_-TYW05vAAj_0V3J8a-KTMImksXMeXe9YOR270TElT_srPu/exec",
        "api_key": ""
    }
}


def load_config() -> dict:
    cfg = _read_json(CONFIG_PATH)
    if not cfg:
        cfg = DEFAULT_CONFIG
        _write_json(CONFIG_PATH, cfg)
    # merge missing keys (lightweight)
    merged = DEFAULT_CONFIG | cfg
    merged["google"] = (DEFAULT_CONFIG["google"] | cfg.get("google", {}))
    return merged


def save_config(cfg: dict) -> None:
    _write_json(CONFIG_PATH, cfg)


def ensure_record_defaults(rec: dict) -> dict:
    """Ensure all expected keys exist."""
    defaults = {
        "id": "",
        "status": "draft",
        "filler_name": "",
        "form_date": date.today().isoformat(),
        "plan_code": "",
        "purpose_desc": "",
        "payment_mode": "employee",  # employee | advance | vendor
        "payee_type": "",  # backward compat / reserved
        "employee_name": "",
        "employee_no": "",
        "vendor_name": "",
        "vendor_address": "",
        "vendor_payee_name": "",
        "is_advance_offset": False,
        "advance_amount": 0,
        "offset_amount": 0,
        "balance_refund_amount": 0,
        "supplement_amount": 0,
        "receipt_no": "",
        "amount_untaxed": 0,
        "tax_amount": 0,
        "amount_total": 0,
        "handler_name": "",
        "project_manager_name": "",
        "dept_manager_name": "",
        "accountant_name": "",
        "attachments": "[]",  # JSON list of relative file paths
        "created_at": "",
        "updated_at": "",
        "submitted_at": "",
    }
    out = defaults | (rec or {})
    # Normalize attachments to JSON string
    out["attachments"] = normalize_attachments_cell(out.get("attachments"))
    return out


def normalize_attachments_cell(cell) -> str:
    """Return a JSON list string."""
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return "[]"
    if isinstance(cell, list):
        return json.dumps(cell, ensure_ascii=False)
    s = str(cell).strip()
    if not s:
        return "[]"
    if s.startswith("["):
        try:
            json.loads(s)
            return s
        except Exception:
            return "[]"
    # legacy single path
    return json.dumps([s], ensure_ascii=False)


def parse_attachments(cell_json: str) -> List[str]:
    try:
        xs = json.loads(cell_json or "[]")
        if isinstance(xs, list):
            return [str(x) for x in xs if str(x).strip()]
        return []
    except Exception:
        return []


def to_float(x, default=0.0) -> float:
    """Coerce common spreadsheet values to float safely."""
    try:
        if x is None:
            return float(default)

        # pandas/openpyxl may yield NaN
        try:
            import math
            if isinstance(x, float) and math.isnan(x):
                return float(default)
        except Exception:
            pass

        if isinstance(x, (int, float)):
            return float(x)

        s = str(x).strip()
        if not s:
            return float(default)
        s = s.replace(",", "").replace("$", "")
        return float(s)
    except Exception:
        return float(default)


def generate_new_id(df: pd.DataFrame, form_date: str) -> str:
    """YYYYMMDD + 3-digit sequence. Sequence is based on local records."""
    d = form_date.replace("-", "")
    prefix = d
    max_seq = 0
    if df is not None and not df.empty and "id" in df.columns:
        for rid in df["id"].astype(str).tolist():
            if rid.startswith(prefix) and len(rid) >= 11:
                tail = rid[len(prefix):len(prefix)+3]
                if tail.isdigit():
                    max_seq = max(max_seq, int(tail))
    return f"{prefix}{max_seq+1:03d}"


def get_local_df() -> pd.DataFrame:
    df = load_local_all(str(LOCAL_XLSX))
    if df is None or df.empty:
        return pd.DataFrame()
    return df


def get_record_by_id(df: pd.DataFrame, rid: str) -> Optional[dict]:
    if df is None or df.empty:
        return None
    m = df[df["id"].astype(str) == str(rid)]
    if m.empty:
        return None
    return ensure_record_defaults(m.iloc[0].to_dict())


def upsert_local_record(rec: dict) -> None:
    rec = ensure_record_defaults(rec)
    rec["updated_at"] = _now_iso()
    if not rec.get("created_at"):
        rec["created_at"] = rec["updated_at"]
    upsert_local(str(LOCAL_XLSX), rec)


def cloud_enabled(cfg: dict) -> bool:
    return cfg.get("backend") == "google"


def cloud_config(cfg: dict) -> dict:
    g = cfg.get("google", {})
    return {
        "spreadsheet_id": parse_sheet_id(g.get("spreadsheet_id", "")),
        "submit_sheet_name": str(g.get("submit_sheet_name", "申請表單")),
        "draft_sheet_name": str(g.get("draft_sheet_name", "草稿列表")),
        "apps_script_url": normalize_apps_script_url(g.get("apps_script_url", "")),
        "api_key": str(g.get("api_key", "")),
    }


def safe_cloud_upsert(cfg: dict, sheet_name: str, rec: dict) -> Tuple[bool, str]:
    g = cloud_config(cfg)
    try:
        cloud_upsert(
            script_url=g["apps_script_url"],
            spreadsheet_id=g["spreadsheet_id"],
            sheet_name=sheet_name,
            payload=rec,
            api_key=g["api_key"],
        )
        return True, "OK"
    except Exception as e:
        return False, str(e)


def safe_cloud_delete(cfg: dict, sheet_name: str, record_id: str) -> Tuple[bool, str]:
    """Best-effort delete a row in cloud sheet.

    Some Google Sheets store numeric IDs as numbers; Apps Script may compare strictly.
    We try once with string id, and if not found and id is digits-only, try again as int.
    """
    g = cloud_config(cfg)
    try:
        deleted = cloud_delete(
            script_url=g["apps_script_url"],
            spreadsheet_id=g["spreadsheet_id"],
            sheet_name=sheet_name,
            record_id=record_id,
            api_key=g["api_key"],
        )
        if (not deleted) and str(record_id).strip().isdigit():
            try:
                deleted = cloud_delete(
                    script_url=g["apps_script_url"],
                    spreadsheet_id=g["spreadsheet_id"],
                    sheet_name=sheet_name,
                    record_id=int(str(record_id).strip()),
                    api_key=g["api_key"],
                )
            except Exception:
                pass
        return True, "deleted" if deleted else "not_found"
    except Exception as e:
        return False, str(e)



def download_local_excel() -> Tuple[bytes, str]:
    """Download an Excel with 2 sheets:
    - 申請表單: submitted / void
    - 草稿列表: draft / deleted
    (local storage remains in apps/data/vouchers.xlsx; this is only the exported file.)
    """
    df = get_local_df()
    try:
        from storage_excel import COLUMNS
    except Exception:
        COLUMNS = list(df.columns) if df is not None else []

    if df is None or df.empty:
        # Still export an empty workbook with headers
        df = pd.DataFrame(columns=COLUMNS)

    for c in COLUMNS:
        if c not in df.columns:
            df[c] = ""

    submit_df = df[df.get("status", "").astype(str).isin(["submitted", "void"])].copy()
    draft_df = df[df.get("status", "").astype(str).isin(["draft", "deleted"])].copy()

    if COLUMNS:
        submit_df = submit_df[COLUMNS]
        draft_df = draft_df[COLUMNS]

    from io import BytesIO
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        submit_df.to_excel(writer, sheet_name="申請表單", index=False)
        draft_df.to_excel(writer, sheet_name="草稿列表", index=False)

    export_name = f"{LOCAL_XLSX.stem}_export.xlsx"
    return bio.getvalue(), export_name


def save_uploaded_files(record_id: str, files: List) -> List[str]:
    """Save uploaded files to data/attachments/<id>/ and return relative paths."""
    folder = ATTACH_DIR / record_id
    folder.mkdir(parents=True, exist_ok=True)
    rel_paths = []
    for f in files:
        name = re.sub(r"[^\w\-.()\[\] ]+", "_", f.name).strip()
        if not name:
            name = "upload"
        target = folder / name
        with open(target, "wb") as out:
            out.write(f.getbuffer())
        rel_paths.append(str(target.relative_to(APP_DIR)))
    return rel_paths


def resolve_attachment_paths(rel_paths: List[str]) -> List[str]:
    out = []
    for rp in rel_paths:
        try:
            p = (APP_DIR / rp).resolve()
            if p.exists():
                out.append(str(p))
        except Exception:
            pass
    return out


# ----------------------------
# UI
# ----------------------------

st.markdown(
    """
    <style>
    .stButton > button { width: 100%; }
    .tight-buttons .stButton > button { padding-top: 0.4rem; padding-bottom: 0.4rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


def sidebar_settings(cfg: dict) -> dict:
    st.sidebar.title("支出報帳系統")
    st.sidebar.caption("資料儲存與雲端寫入設定（會記錄在 data/config.json）")

    backend = "google"

    g = cloud_config(cfg)

    with st.sidebar.expander("Google Sheet 設定", expanded=True):
        spreadsheet_id = st.text_input("Sheet ID（或貼整個網址）", value=g["spreadsheet_id"], key="cfg_sheet_id")
        apps_script_url = st.text_input("Apps Script URL（或部署 ID AKfy...）", value=g["apps_script_url"], key="cfg_apps_url")
        submit_sheet_name = st.text_input("送出工作表名稱", value=g["submit_sheet_name"], key="cfg_submit_sheet")
        draft_sheet_name = st.text_input("草稿工作表名稱", value=g["draft_sheet_name"], key="cfg_draft_sheet")
        api_key = st.text_input("API Key（可留空）", value=g["api_key"], type="password", key="cfg_api_key")

    if st.sidebar.button("儲存設定", use_container_width=True):
        new_cfg = dict(cfg)
        new_cfg["backend"] = "google"
        new_cfg["google"] = {
            "spreadsheet_id": parse_sheet_id(spreadsheet_id),
            "submit_sheet_name": submit_sheet_name.strip() or "申請表單",
            "draft_sheet_name": draft_sheet_name.strip() or "草稿列表",
            "apps_script_url": normalize_apps_script_url(apps_script_url),
            "api_key": api_key.strip(),
        }
        save_config(new_cfg)
        st.sidebar.success("已儲存")
        return new_cfg

    st.sidebar.divider()
    st.sidebar.caption("快速操作")
    xls_bytes, xls_name = download_local_excel()
    st.sidebar.download_button("下載 Excel", data=xls_bytes, file_name=xls_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    open_url = f"https://docs.google.com/spreadsheets/d/{g['spreadsheet_id']}" if g.get("spreadsheet_id") else ""
    if open_url:
        st.sidebar.link_button("開啟雲端表單", open_url, use_container_width=True)
    else:
        st.sidebar.button("開啟雲端表單", disabled=True, use_container_width=True)

    return cfg



cfg = sidebar_settings(load_config())

# Top navigation (moved from sidebar)
def _top_nav_changed():
    st.session_state["page"] = st.session_state.get("top_nav_page", "list")

current_page = st.session_state.get("page", "list")
opts = ["list", "new", "drafts"]
default_idx = opts.index(current_page) if current_page in opts else 0
st.radio(
    "功能選單",
    options=opts,
    format_func=lambda x: {"list":"表單列表/查詢", "new":"新增表單", "drafts":"草稿列表"}[x],
    key="top_nav_page",
    index=default_idx,
    horizontal=True,
    on_change=_top_nav_changed,
)

# ----------------------------
# Pages
# ----------------------------
def render_row_actions(rows_df: pd.DataFrame, scope: str):
    """Table-like rows with right-side action buttons (編輯/下載/送出/作廢或刪除)."""
    if rows_df is None or rows_df.empty:
        return

    # Header row (to mimic a table action column)
    h1,h2,h3,h4 = st.columns([1,1,1,1])
    with h1: st.markdown("**編輯**")
    with h2: st.markdown("**下載**")
    with h3: st.markdown("**送出**")
    with h4: st.markdown("**作廢/刪除**")

    for _, r in rows_df.iterrows():
        rid = str(r.get("id",""))
        status = str(r.get("status",""))
        c1,c2,c3,c4 = st.columns([1,1,1,1])
        with c1:
            if st.button("編輯", key=f"{scope}_edit_{rid}", use_container_width=True):
                st.session_state["current_id"]=rid
                st.session_state["page"]="edit"
                st.rerun()
        with c2:
            if st.button("下載", key=f"{scope}_dl_{rid}", use_container_width=True):
                st.session_state["current_id"]=rid
                st.session_state["page"]="view"
                st.session_state["auto_download_pdf"]=True
                st.rerun()
        with c3:
            if st.button("送出", key=f"{scope}_submit_{rid}", use_container_width=True):
                df_all=get_local_df(); rec=get_record_by_id(df_all,rid)
                if rec:
                    rec["status"]="submitted"
                    rec["submitted_at"]=_now_iso()
                    upsert_local_record(rec)
                    if cloud_enabled(cfg):
                        g=cloud_config(cfg)
                        safe_cloud_upsert(cfg,g["submit_sheet_name"],rec)
                        safe_cloud_delete(cfg,g["draft_sheet_name"],rid)
                    st.success(f"已送出 {rid}")
                    st.rerun()
        with c4:
            if status=="submitted":
                if st.button("作廢", key=f"{scope}_void_{rid}", use_container_width=True):
                    df_all=get_local_df(); rec=get_record_by_id(df_all,rid)
                    if rec:
                        rec["status"]="void"
                        upsert_local_record(rec)
                        if cloud_enabled(cfg):
                            g=cloud_config(cfg)
                            safe_cloud_upsert(cfg,g["submit_sheet_name"],rec)
                        st.success(f"已作廢 {rid}")
                        st.rerun()
            else:
                if st.button("刪除", key=f"{scope}_del_{rid}", use_container_width=True):
                    delete_local_record(str(LOCAL_XLSX), rid)
                    if cloud_enabled(cfg):
                        g=cloud_config(cfg)
                        safe_cloud_delete(cfg,g["draft_sheet_name"],rid)
                    st.success(f"已刪除 {rid}")
                    st.rerun()


def page_list():
    st.header("表單列表/查詢")

    df = get_local_df()
    if df.empty:
        st.info("目前本機尚無資料。請到左側選『新增表單』開始。")
        return

    df = df.copy()
    # Normalize columns to avoid KeyError when sources use different column names
    if 'id' not in df.columns: df['id'] = ''
    if 'status' not in df.columns: df['status'] = ''
    if 'form_date' not in df.columns:
        for _alt in ['date','日期','填寫日期','送出日期']:
            if _alt in df.columns:
                df['form_date'] = df[_alt]
                break
        else:
            df['form_date'] = ''
    for col in ["amount_untaxed","tax_amount","amount_total","advance_amount","offset_amount","balance_refund_amount","supplement_amount"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # This page shows only submitted/void
    base_df = df[df.get("status","").astype(str).isin(["submitted","void"])].copy()
    if base_df.empty:
        st.info("目前尚無已送出或已作廢資料（submitted/void）。")
        return

    # Default YM range
    start_default = str(base_df["form_date"].astype(str).min())[:7] if "form_date" in base_df.columns else date.today().strftime("%Y-%m")
    end_default = str(base_df["form_date"].astype(str).max())[:7] if "form_date" in base_df.columns else date.today().strftime("%Y-%m")

    # Guard against invalid defaults (e.g., '0 20...' from Series str)
    if not re.match(r'^\d{4}-\d{2}$', str(start_default).strip()):
        start_default = '2026-01'
    if not re.match(r'^\d{4}-\d{2}$', str(end_default).strip()):
        end_default = date.today().strftime('%Y-%m')

    # User-requested default
    start_default = '2026-01'


    # init session state (so reset works)
    st.session_state.setdefault("list_status", "(全部)")
    st.session_state.setdefault("list_filler", "")
    st.session_state.setdefault("list_plan", "")
    st.session_state.setdefault("list_id", "")
    st.session_state.setdefault("list_start", start_default)
    st.session_state.setdefault("list_end", end_default)

    # If session has legacy invalid values, normalize them now so UI shows correct defaults
    if not re.match(r'^\d{4}-\d{2}$', str(st.session_state.get('list_start','')).strip()):
        st.session_state['list_start'] = '2026-01'
    if not re.match(r'^\d{4}-\d{2}$', str(st.session_state.get('list_end','')).strip()):
        st.session_state['list_end'] = end_default

    # reset button (same spot as other pages)
    if st.button("重設篩選", key="list_reset_filters"):
        st.session_state["list_status"] = "(全部)"
        st.session_state["list_filler"] = ""
        st.session_state["list_plan"] = ""
        st.session_state["list_id"] = ""
        st.session_state["list_start"] = start_default
        st.session_state["list_end"] = end_default
        st.rerun()

    c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2, 1.2])
    with c1:
        status = st.selectbox("狀態", options=["(全部)", "submitted", "void"], key="list_status")
    with c2:
        filler = st.text_input("填表人包含", key="list_filler")
    with c3:
        plan = st.text_input("計畫編號包含", key="list_plan")
    with c4:
        rid = st.text_input("表單ID", key="list_id")

    c5, c6 = st.columns([1.2, 1.2])
    with c5:
        start_ym = st.text_input("起始年月(YYYY-MM)", key="list_start")
    with c6:
        end_ym = st.text_input("結束年月(YYYY-MM)", key="list_end")

    # Normalize YM filters (avoid invalid strings breaking comparisons)
    if not re.match(r'^\d{4}-\d{2}$', str(start_ym).strip()):
        start_ym = st.session_state['list_start'] = '2026-01'
    if not re.match(r'^\d{4}-\d{2}$', str(end_ym).strip()):
        end_ym = st.session_state['list_end'] = date.today().strftime('%Y-%m')

    view = base_df.copy()
    if status != "(全部)":
        view = view[view["status"].astype(str) == status]
    if filler:
        view = view[view.get("filler_name","").astype(str).str.contains(filler, na=False)]
    if plan:
        view = view[view.get("plan_code","").astype(str).str.contains(plan, na=False)]
    if rid:
        view = view[view["id"].astype(str).str.contains(rid, na=False)]

    def _in_ym(d):
        d = str(d)
        return len(d) >= 7 and start_ym <= d[:7] <= end_ym

    if "form_date" in view.columns:
        view = view[view["form_date"].astype(str).apply(_in_ym)]

    view = view.sort_values(by=["form_date","id"], ascending=[False, False])

    view_show = view.copy()
    view_show["狀態標籤"] = view_show.get("status","").astype(str).map(
        {"draft":"⚪ draft","submitted":"🟢 submitted","void":"🔴 void","deleted":"⚫ deleted"}
    ).fillna(view_show.get("status","").astype(str))
    view_show["事由說明"] = view_show.get("purpose_desc","").astype(str).apply(lambda x: short_text(x, 10))
    view_show["_總金額"] = pd.to_numeric(view_show.get("amount_total",0), errors="coerce").fillna(0).apply(lambda v: f"{v:,.1f}")
    st.dataframe(
        view_show[[
            "id","狀態標籤","form_date","filler_name","plan_code","payment_mode","_總金額","事由說明","updated_at"
        ]].rename(columns={
            "id":"表單ID","狀態標籤":"狀態標籤","form_date":"日期","filler_name":"填表人","plan_code":"計畫編號",
            "payment_mode":"付款對象","_總金額":"總金額","updated_at":"更新時間"
        }),
        use_container_width=True,
        hide_index=True,
    )

    total_text = (
        f"區間合計：未稅 {pd.to_numeric(view.get('amount_untaxed',0), errors='coerce').fillna(0).sum():,.1f} / "
        f"稅金 {pd.to_numeric(view.get('tax_amount',0), errors='coerce').fillna(0).sum():,.1f} / "
        f"總金額 {pd.to_numeric(view.get('amount_total',0), errors='coerce').fillna(0).sum():,.1f} / "
        f"筆數 {len(view)}"
    )
    st.markdown(f"<div style='font-size:2.1rem;font-weight:800;line-height:1.2;margin-top:0.25rem'>{total_text}</div>", unsafe_allow_html=True)

    st.divider()
    st.subheader("快速操作（先選表單ID，再執行動作）")

    if view.empty:
        st.info("目前篩選後無資料。")
        return

    select_id = st.selectbox(
        "表單ID",
        options=[""] + view["id"].astype(str).tolist(),
        format_func=lambda x: x if not x else (
            f"{x}｜{str(view[view['id'].astype(str)==str(x)].iloc[0].get('status',''))}"
            f"｜{str(view[view['id'].astype(str)==str(x)].iloc[0].get('filler_name',''))}"
            f"｜計畫:{str(view[view['id'].astype(str)==str(x)].iloc[0].get('plan_code',''))}"
            f"｜金額:{pd.to_numeric(view[view['id'].astype(str)==str(x)].iloc[0].get('amount_total',0), errors='coerce'):.1f}"
            f"｜事由:{short_text(view[view['id'].astype(str)==str(x)].iloc[0].get('purpose_desc',''),10)}"
        ),
        key="list_action_pick"
    )

    rr = get_record_by_id(df, select_id) if select_id else None
    cA, cB, cC, cD = st.columns(4)

    with cA:
        if st.button("編輯", disabled=not bool(select_id), use_container_width=True, key="list_action_edit"):
            st.session_state["current_id"] = select_id
            st.session_state["page"] = "edit"
            st.rerun()

    with cB:
        if st.button("下載", disabled=not bool(select_id), use_container_width=True, key="list_action_dl"):
            st.session_state["current_id"] = select_id
            st.session_state["page"] = "view"
            st.session_state["auto_download_pdf"] = True
            st.rerun()

    with cC:
        # list page doesn't handle draft submit; keep button disabled for safety
        st.button("送出", disabled=True, use_container_width=True, key="list_action_submit_disabled")

    with cD:
        can_void = bool(rr) and str(rr.get("status","")) == "submitted"
        if st.button("作廢", disabled=not can_void, use_container_width=True, key="list_action_void"):
            rec = get_record_by_id(get_local_df(), select_id)
            if rec:
                rec["status"] = "void"
                upsert_local_record(rec)
                if cloud_enabled(cfg):
                    g = cloud_config(cfg)
                    safe_cloud_upsert(cfg, g["submit_sheet_name"], rec)
                st.success(f"已作廢 {select_id}")
            st.rerun()

def page_new():
    st.header("新增表單")
    df = get_local_df()
    form_date = st.date_input("表單日期", value=date.today())
    if st.button("建立新表單", use_container_width=True):
        rid = generate_new_id(df, form_date.isoformat())
        rec = ensure_record_defaults({"id": rid, "form_date": form_date.isoformat(), "status": "draft"})
        rec["created_at"] = _now_iso()
        rec["updated_at"] = rec["created_at"]
        upsert_local_record(rec)
        st.session_state["current_id"] = rid
        st.session_state["page"] = "edit"
        st.rerun()


def page_drafts():
    st.header("草稿列表")
    df = get_local_df()
    if df.empty:
        st.info("目前沒有資料。")
        return

    df = df.copy()
    if 'id' not in df.columns: df['id'] = ''
    if 'status' not in df.columns: df['status'] = ''
    if 'form_date' not in df.columns:
        for _alt in ['date','日期','填寫日期','送出日期']:
            if _alt in df.columns:
                df['form_date'] = df[_alt]
                break
        else:
            df['form_date'] = ''

    # only draft/deleted for this page
    base_df = df[df.get("status", "").astype(str).isin(["draft", "deleted"])].copy()
    if base_df.empty:
        st.info("目前沒有草稿或已刪除資料（draft/deleted）。")
        return

    # default YM range
    start_default = str(base_df['form_date'].astype(str).min())[:7] if 'form_date' in base_df.columns else ''
    end_default = str(base_df['form_date'].astype(str).max())[:7] if 'form_date' in base_df.columns else ''
    if not re.match(r'^\d{4}-\d{2}$', str(start_default).strip()):
        start_default = '2026-01'
    if not re.match(r'^\d{4}-\d{2}$', str(end_default).strip()):
        end_default = date.today().strftime('%Y-%m')

    # User-requested default
    start_default = '2026-01'


    # init session state (so reset works)
    st.session_state.setdefault("draft_status", "(全部)")
    st.session_state.setdefault("draft_filler", "")
    st.session_state.setdefault("draft_plan", "")
    st.session_state.setdefault("draft_id", "")
    st.session_state.setdefault("draft_start", start_default)
    st.session_state.setdefault("draft_end", end_default)

    if not re.match(r'^\d{4}-\d{2}$', str(st.session_state.get('draft_start','')).strip()):
        st.session_state['draft_start'] = '2026-01'
    if not re.match(r'^\d{4}-\d{2}$', str(st.session_state.get('draft_end','')).strip()):
        st.session_state['draft_end'] = end_default

    if st.button("重設篩選", key="draft_reset_filters"):
        st.session_state["draft_status"] = "(全部)"
        st.session_state["draft_filler"] = ""
        st.session_state["draft_plan"] = ""
        st.session_state["draft_id"] = ""
        st.session_state["draft_start"] = start_default
        st.session_state["draft_end"] = end_default
        st.rerun()

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        status = st.selectbox("狀態", options=["(全部)", "draft", "deleted"], key="draft_status")
    with c2:
        filler = st.text_input("填表人包含", key="draft_filler")
    with c3:
        plan = st.text_input("計畫編號包含", key="draft_plan")
    with c4:
        qid = st.text_input("表單ID", key="draft_id")
    c5, c6 = st.columns(2)
    with c5:
        start_ym = st.text_input("起始年月(YYYY-MM)", key="draft_start")
    with c6:
        end_ym = st.text_input("結束年月(YYYY-MM)", key="draft_end")

    if not re.match(r'^\d{4}-\d{2}$', str(start_ym).strip()):
        start_ym = st.session_state['draft_start'] = '2026-01'
    if not re.match(r'^\d{4}-\d{2}$', str(end_ym).strip()):
        end_ym = st.session_state['draft_end'] = date.today().strftime('%Y-%m')

    view = base_df.copy()
    if status != "(全部)":
        view = view[view.get("status", "").astype(str) == status]
    if filler:
        view = view[view.get("filler_name", "").astype(str).str.contains(filler, na=False)]
    if plan:
        view = view[view.get("plan_code", "").astype(str).str.contains(plan, na=False)]
    if qid:
        view = view[view.get("id", "").astype(str).str.contains(qid, na=False)]
    if start_ym and "form_date" in view.columns:
        view = view[view.get("form_date", "").astype(str) >= start_ym + "-01"]
    if end_ym and "form_date" in view.columns:
        view = view[view.get("form_date", "").astype(str) <= end_ym + "-31"]

    view = view.sort_values(by=["form_date", "id"], ascending=[False, False])

    vshow = view.copy()
    vshow["狀態標籤"] = vshow.get("status", "").astype(str).map(
        {"draft": "⚪ draft", "submitted": "🟢 submitted", "void": "🔴 void", "deleted": "⚫ deleted"}
    ).fillna(vshow.get("status", "").astype(str))
    vshow["事由說明"] = vshow.get("purpose_desc", "").astype(str).apply(lambda x: short_text(x, 10))
    vshow["_總金額"] = pd.to_numeric(vshow.get("amount_total", 0), errors="coerce").fillna(0).apply(lambda v: f"{v:,.1f}")

    st.dataframe(
        vshow[["id", "狀態標籤", "form_date", "filler_name", "plan_code", "payment_mode", "_總金額", "事由說明", "updated_at"]].rename(
            columns={
                "id": "表單ID",
                "form_date": "日期",
                "filler_name": "填表人",
                "plan_code": "計畫編號",
                "payment_mode": "付款對象",
                "_總金額": "總金額",
                "updated_at": "更新時間",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    total_text = (
        f"區間合計：未稅 {pd.to_numeric(view.get('amount_untaxed',0),errors='coerce').fillna(0).sum():,.1f} / "
        f"稅金 {pd.to_numeric(view.get('tax_amount',0),errors='coerce').fillna(0).sum():,.1f} / "
        f"總金額 {pd.to_numeric(view.get('amount_total',0),errors='coerce').fillna(0).sum():,.1f} / "
        f"筆數 {len(view)}"
    )
    st.markdown(f"<div style='font-size:2.1rem;font-weight:800;line-height:1.2;margin-top:0.25rem'>{total_text}</div>", unsafe_allow_html=True)

    st.divider()
    st.subheader("快速操作（先選表單ID，再執行動作）")

    if view.empty:
        st.info("目前篩選後無資料。")
        return

    select_id = st.selectbox(
        "表單ID",
        options=[""] + view["id"].astype(str).tolist(),
        format_func=lambda x: x if not x else (
            f"{x}｜{str(view[view['id'].astype(str)==str(x)].iloc[0].get('status',''))}"
            f"｜{str(view[view['id'].astype(str)==str(x)].iloc[0].get('filler_name',''))}"
            f"｜計畫:{str(view[view['id'].astype(str)==str(x)].iloc[0].get('plan_code',''))}"
            f"｜金額:{pd.to_numeric(view[view['id'].astype(str)==str(x)].iloc[0].get('amount_total',0), errors='coerce'):.1f}"
            f"｜事由:{short_text(view[view['id'].astype(str)==str(x)].iloc[0].get('purpose_desc',''),10)}"
        ),
        key="draft_action_pick",
    )

    rr = get_record_by_id(df, select_id) if select_id else None
    cA, cB, cC, cD = st.columns(4)

    with cA:
        if st.button("編輯", disabled=not bool(select_id), use_container_width=True, key="draft_action_edit"):
            st.session_state["current_id"] = select_id
            st.session_state["page"] = "edit"
            st.rerun()

    with cB:
        if st.button("下載", disabled=not bool(select_id), use_container_width=True, key="draft_action_dl"):
            st.session_state["current_id"] = select_id
            st.session_state["page"] = "view"
            st.session_state["auto_download_pdf"] = True
            st.rerun()

    with cC:
        can_submit = bool(rr) and str(rr.get("status", "")) == "draft"
        if st.button("送出", disabled=not can_submit, use_container_width=True, key="draft_action_submit"):
            rec = rr
            rec["status"] = "submitted"
            rec["submitted_at"] = _now_iso()
            upsert_local_record(rec)
            if cloud_enabled(cfg):
                g = cloud_config(cfg)
                safe_cloud_upsert(cfg, g["submit_sheet_name"], rec)
                safe_cloud_delete(cfg, g["draft_sheet_name"], select_id)
            st.success(f"已送出 {select_id}")
            st.rerun()

    with cD:
        can_del = bool(rr) and str(rr.get("status", "")) in ("draft", "deleted")
        if st.button("刪除", disabled=not can_del, use_container_width=True, key="draft_action_delete"):
            rec = rr
            rec["status"] = "deleted"
            upsert_local_record(rec)
            if cloud_enabled(cfg):
                g = cloud_config(cfg)
                safe_cloud_upsert(cfg, g["draft_sheet_name"], rec)
            st.success(f"已標記刪除 {select_id}")
            st.rerun()


def page_edit():
    rid = st.session_state.get("current_id", "")
    st.header(f"編輯表單：{rid}")

    df = get_local_df()
    rec = get_record_by_id(df, rid)
    if not rec:
        st.error("找不到此表單。")
        return

    # Initialize session state fields for this record
    key_prefix = f"rec_{rid}_"
    for k, v in rec.items():
        st.session_state.setdefault(key_prefix + k, v)

    # Normalize legacy / spreadsheet values (prevents Streamlit widget type errors)
    # payment_mode may be stored as Chinese labels in older files
    pm_key = key_prefix + "payment_mode"
    pm_map = {
        "員工姓名": "employee",
        "借支沖抵": "advance",
        "廠商付款": "vendor",
        "逕付廠商": "vendor",
    }
    if pm_key in st.session_state:
        pm_val = str(st.session_state.get(pm_key, "employee")).strip()
        st.session_state[pm_key] = pm_map.get(pm_val, pm_val)
        if st.session_state[pm_key] not in ("employee", "advance", "vendor"):
            st.session_state[pm_key] = "employee"

    # Coerce numeric widget state to float (number_input expects numeric session state)
    for nk in [
        "advance_amount",
        "offset_amount",
        "balance_refund_amount",
        "supplement_amount",
        "amount_untaxed",
        "tax_amount",
        "amount_total",
    ]:
        k = key_prefix + nk
        if k in st.session_state:
            st.session_state[k] = to_float(st.session_state.get(k))

    # --- Form fields ---
    left, right = st.columns([1.4, 1.0])

    with left:
        st.subheader("基本資訊")

        form_date = st.date_input("表單日期", value=date.fromisoformat(st.session_state[key_prefix+"form_date"]), key=key_prefix+"form_date_ui")
        st.session_state[key_prefix+"form_date"] = form_date.isoformat()

        st.text_input("填表人", value=st.session_state[key_prefix+"filler_name"], key=key_prefix+"filler_name")
        st.text_input("計畫編號", value=st.session_state[key_prefix+"plan_code"], key=key_prefix+"plan_code")
        st.text_area("事由說明", value=st.session_state[key_prefix+"purpose_desc"], key=key_prefix+"purpose_desc", height=90)

        st.subheader("付款對象（互斥三選一）")
        payment_mode = st.radio(
            "付款對象",
            options=["employee","advance","vendor"],
            format_func=lambda x: {"employee":"員工姓名","advance":"借支沖抵","vendor":"廠商付款"}[x],
            horizontal=True,
            key=key_prefix+"payment_mode",
        )

        if payment_mode == "employee":
            st.text_input("員工姓名", value=st.session_state[key_prefix+"employee_name"], key=key_prefix+"employee_name")
            st.text_input("員工編號", value=st.session_state[key_prefix+"employee_no"], key=key_prefix+"employee_no")
        elif payment_mode == "advance":
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.number_input("借支金額", min_value=0.0, value=float(to_float(st.session_state[key_prefix+"advance_amount"])), step=0.1, key=key_prefix+"advance_amount")
            with c2:
                st.number_input("沖抵金額", min_value=0.0, value=float(to_float(st.session_state[key_prefix+"offset_amount"])), step=0.1, key=key_prefix+"offset_amount")
            with c3:
                st.number_input("結餘繳回", min_value=0.0, value=float(to_float(st.session_state[key_prefix+"balance_refund_amount"])), step=0.1, key=key_prefix+"balance_refund_amount")
            with c4:
                st.number_input("不足補付", min_value=0.0, value=float(to_float(st.session_state[key_prefix+"supplement_amount"])), step=0.1, key=key_prefix+"supplement_amount")
        else:
            st.text_input("廠商名稱", value=st.session_state[key_prefix+"vendor_name"], key=key_prefix+"vendor_name")
            st.text_input("廠商地址", value=st.session_state[key_prefix+"vendor_address"], key=key_prefix+"vendor_address")
            st.text_input("廠商收款人（若不同）", value=st.session_state[key_prefix+"vendor_payee_name"], key=key_prefix+"vendor_payee_name")

        st.subheader("金額與憑證")
        c1, c2, c3, c4 = st.columns([1,1,1,1])
        with c1:
            st.text_input("憑證號碼", value=st.session_state[key_prefix+"receipt_no"], key=key_prefix+"receipt_no")
        with c2:
            st.number_input("未稅金額", min_value=0.0, value=float(to_float(st.session_state[key_prefix+"amount_untaxed"])), step=0.1, key=key_prefix+"amount_untaxed")
        with c3:
            st.number_input("稅金", min_value=0.0, value=float(to_float(st.session_state[key_prefix+"tax_amount"])), step=0.1, key=key_prefix+"tax_amount")
        amt_total = to_float(st.session_state.get(key_prefix+"amount_untaxed",0))+to_float(st.session_state.get(key_prefix+"tax_amount",0))
        st.session_state[key_prefix+"amount_total"] = amt_total
        with c4:
            st.number_input("總金額", min_value=0.0, value=float(amt_total), step=0.1, key=key_prefix+"amount_total", disabled=True)

        st.subheader("簽核欄位")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.text_input("經辦人", value=st.session_state[key_prefix+"handler_name"], key=key_prefix+"handler_name")
        with c2:
            st.text_input("專案經理", value=st.session_state[key_prefix+"project_manager_name"], key=key_prefix+"project_manager_name")
        with c3:
            st.text_input("部門主管", value=st.session_state[key_prefix+"dept_manager_name"], key=key_prefix+"dept_manager_name")
        with c4:
            st.text_input("會計", value=st.session_state[key_prefix+"accountant_name"], key=key_prefix+"accountant_name")

    with right:
        st.subheader("附件（收據/發票/照片/PDF）")
        existing = parse_attachments(st.session_state[key_prefix+"attachments"])
        existing_abs = resolve_attachment_paths(existing)
        if existing_abs:
            st.caption("已儲存附件：")
            for p in existing_abs:
                st.write(f"- {Path(p).name}")

        uploaded = st.file_uploader(
            "新增附件（可多選）",
            type=["pdf","png","jpg","jpeg","webp"],
            accept_multiple_files=True,
            key=key_prefix+"uploader",
        )

        if uploaded:
            new_rel = save_uploaded_files(rid, uploaded)
            # Deduplicate by casting to dict (ordered) mapping path -> True
            merged_dict = {x: True for x in existing}
            for p in new_rel: merged_dict[p] = True
            merged = list(merged_dict.keys())
            st.session_state[key_prefix+"attachments"] = json.dumps(merged, ensure_ascii=False)
            st.success(f"已加入 {len(new_rel)} 個附件（將在儲存/送出時一起寫入本機）")

        st.divider()
        st.caption("PDF 匯出（含附件合併）")

        if "last_pdf_bytes" not in st.session_state:
            st.session_state["last_pdf_bytes"] = None

        current = collect_record_from_state(key_prefix, rid)
        paths = resolve_attachment_paths(parse_attachments(current["attachments"]))
        try:
            st.session_state["last_pdf_bytes"] = build_pdf_bytes(current, str(BG_IMAGE), attachment_paths=paths)
            st.download_button("下載 PDF",data=st.session_state["last_pdf_bytes"],file_name=f"支出報帳_{rid}.pdf",mime="application/pdf",use_container_width=True,key=f"dlpdf_{rid}")
        except Exception as e:
            st.warning(f"PDF 產生失敗：{e}")

    # --- Action buttons (one line) ---
    st.divider()
    st.markdown('<div class="tight-buttons">', unsafe_allow_html=True)
    b1, b2, b3, b4, b5 = st.columns([1, 1, 1, 1,1])
    is_submitted = str(rec.get("status","")) == "submitted"
    with b1:
        save_draft_clicked = st.button("儲存草稿（本機+雲端草稿）", disabled=is_submitted, use_container_width=True, key=f"save_{rid}")
    with b2:
        submit_clicked = st.button("送出（本機+雲端送出）", use_container_width=True, key=f"submit_{rid}")
    with b3:
        view_clicked = st.button("只查看（不儲存）", use_container_width=True, key=f"view_{rid}")
    with b4:
        void_clicked = st.button("作廢", disabled=not is_submitted, use_container_width=True, key=f"void_{rid}")
    with b5:
        back_clicked = st.button("返回列表", use_container_width=True, key=f"back_{rid}")
    st.markdown("</div>", unsafe_allow_html=True)

    if back_clicked:
        st.session_state["page"] = "list"
        st.rerun()

    if view_clicked:
        st.session_state["page"] = "view"
        st.rerun()

    if void_clicked:
        current = collect_record_from_state(key_prefix, rid)
        current["status"] = "void"
        upsert_local_record(current)
        if cloud_enabled(cfg):
            g = cloud_config(cfg)
            safe_cloud_upsert(cfg, g["submit_sheet_name"], current)
            safe_cloud_delete(cfg, g["draft_sheet_name"], rid)
        st.success("此筆已作廢")
        st.session_state["page"]="view"
        st.rerun()

    if save_draft_clicked or submit_clicked:
        current = collect_record_from_state(key_prefix, rid)

        # Normalize mode-related fields
        pm = current.get("payment_mode", "employee")
        current["payee_type"] = pm  # keep in sync
        current["is_advance_offset"] = (pm == "advance")

        if submit_clicked:
            current["status"] = "submitted"
            current["submitted_at"] = _now_iso()
        else:
            current["status"] = "draft"
            current["submitted_at"] = ""

        # Save local first
        upsert_local_record(current)

        cloud_msgs = []
        if cloud_enabled(cfg):
            g = cloud_config(cfg)
            if not g["spreadsheet_id"] or not g["apps_script_url"]:
                cloud_msgs.append("雲端設定不完整（請到左側儲存設定）")
            else:
                if submit_clicked:
                    ok, msg = safe_cloud_upsert(cfg, g["submit_sheet_name"], current)
                    cloud_msgs.append(f"雲端送出：{'成功' if ok else '失敗'}（{msg}）")
                    # best-effort delete from draft sheet
                    ok2, msg2 = safe_cloud_delete(cfg, g["draft_sheet_name"], rid)
                    cloud_msgs.append(f"雲端草稿移除：{'成功' if (ok2 or msg2 in ['not_found','deleted']) else '失敗'}（{msg2}）")
                else:
                    ok, msg = safe_cloud_upsert(cfg, g["draft_sheet_name"], current)
                    cloud_msgs.append(f"雲端草稿：{'成功' if ok else '失敗'}（{msg}）")

        st.session_state["cloud_msgs"] = cloud_msgs
        st.session_state["current_id"] = rid
        st.session_state["page"] = "view"
        st.rerun()


def collect_record_from_state(key_prefix: str, rid: str) -> dict:
    keys = [
        "id","status","filler_name","form_date","plan_code","purpose_desc",
        "payment_mode","payee_type",
        "employee_name","employee_no",
        "vendor_name","vendor_address","vendor_payee_name",
        "is_advance_offset","advance_amount","offset_amount","balance_refund_amount","supplement_amount",
        "receipt_no","amount_untaxed","tax_amount","amount_total",
        "handler_name","project_manager_name","dept_manager_name","accountant_name",
        "attachments","created_at","updated_at","submitted_at",
    ]
    rec = {}
    for k in keys:
        if k == "id":
            rec[k] = rid
            continue
        v = st.session_state.get(key_prefix + k, "")
        rec[k] = v

    # Numeric normalize
    for k in ["advance_amount","offset_amount","balance_refund_amount","supplement_amount","amount_untaxed","tax_amount","amount_total"]:
        rec[k] = to_float(rec.get(k, 0.0), 0.0)

    # Dates normalize
    rec["form_date"] = str(st.session_state.get(key_prefix+"form_date", date.today().isoformat()))
    rec["created_at"] = str(rec.get("created_at") or "")
    rec["updated_at"] = str(rec.get("updated_at") or "")
    rec["submitted_at"] = str(rec.get("submitted_at") or "")
    rec["amount_total"] = to_float(rec.get("amount_untaxed",0))+to_float(rec.get("tax_amount",0))
    rec["attachments"] = normalize_attachments_cell(rec.get("attachments"))
    return ensure_record_defaults(rec)


def page_view():
    rid = st.session_state.get("current_id", "")
    st.header(f"表單內容：{rid}")

    df = get_local_df()
    rec = get_record_by_id(df, rid)
    if not rec:
        st.error("找不到此表單。")
        return

    msgs = st.session_state.pop("cloud_msgs", [])
    if msgs:
        st.info(" / ".join(msgs))

    # Pretty summary
    st.subheader("摘要")
    summary = {
        "表單ID": rec["id"],
        "狀態": rec["status"],
        "日期": rec["form_date"],
        "填表人": rec.get("filler_name",""),
        "計畫編號": rec.get("plan_code",""),
        "付款對象": rec.get("payment_mode",""),
        "總金額": rec.get("amount_total",0),
        "更新時間": rec.get("updated_at",""),
        "送出時間": rec.get("submitted_at",""),
    }
    st.json(summary, expanded=True)

    st.subheader("完整欄位")
    st.json(rec, expanded=False)

    st.divider()
    c1, c2, c3 = st.columns([1,1,1])
    with c1:
        if st.button("編輯", use_container_width=True):
            st.session_state["page"] = "edit"
            st.rerun()
    with c2:
        if st.button("返回列表", use_container_width=True):
            st.session_state["page"] = "list"
            st.rerun()
    with c3:
        paths = resolve_attachment_paths(parse_attachments(rec["attachments"]))
        try:
            pdf_bytes = build_pdf_bytes(rec, str(BG_IMAGE), attachment_paths=paths)
            st.download_button("下載 PDF", data=pdf_bytes, file_name=f"支出報帳_{rid}.pdf", mime="application/pdf", use_container_width=True)
            if st.session_state.pop("auto_download_pdf", False):
                st.info("請點擊上方『下載 PDF』完成下載。")
        except Exception as e:
            st.warning(f"PDF 產生失敗：{e}")


# Router
page = st.session_state.get("page", "list")
if page == "list":
    page_list()
elif page == "new":
    page_new()
elif page == "drafts":
    page_drafts()
elif page == "edit":
    page_edit()
elif page == "view":
    page_view()
else:
    page_list()

def mark_deleted(record_id: str, cfg: dict):
    df = get_local_df()
    r = get_record_by_id(df, record_id)
    if not r:
        return False, 'not_found'
    r['status'] = 'deleted'
    r['updated_at'] = _now_iso()
    upsert_local_record(r)
    if cloud_enabled(cfg):
        g = cloud_config(cfg)
        safe_cloud_upsert(cfg, g['draft_sheet_name'], r)
    return True, 'marked_deleted'

