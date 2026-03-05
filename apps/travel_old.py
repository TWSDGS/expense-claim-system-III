# apps/travel.py
from __future__ import annotations

import json
import os
import re
import shutil
from datetime import datetime, date, time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from storage_excel_travel import load_all_travel, upsert_travel_record, delete_travel_record
from storage_apps_script import (
    AppsScriptError,
    list_records as cloud_list,
    upsert_record as cloud_upsert,
    delete_record as cloud_delete,
    get_user_profile as cloud_get_user_profile,
    upsert_user_profile as cloud_upsert_user_profile,
    send_pdf_email as cloud_send_pdf_email,
)
from pdf_gen_travel import build_pdf_bytes

APP_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = APP_DIR / "data"
ATTACH_DIR = DATA_DIR / "attachments"
CONFIG_PATH = DATA_DIR / "travel_config.json"
LOCAL_XLSX = DATA_DIR / "travel_vouchers.xlsx"

DATA_DIR.mkdir(exist_ok=True, parents=True)
ATTACH_DIR.mkdir(exist_ok=True, parents=True)


def inject_travel_ui_css() -> None:
    """Soft card-like UI for travel form (no external deps)."""
    st.markdown(
        """
<style>
/* section header */
.travel-title{font-size:44px;font-weight:800;letter-spacing:1px;margin:0 0 6px 0;}
.travel-sub{color:#6b7280;margin:0 0 18px 0;font-size:15px;}
/* subtle card */
.travel-card{
  background: #ffffff;
  border: 1px solid rgba(148,163,184,.35);
  border-radius: 16px;
  padding: 18px 18px 8px 18px;
  box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
  margin-bottom: 16px;
}
.travel-card .h{
  font-size: 22px;
  font-weight: 800;
  margin-bottom: 8px;
}
.travel-note{
  font-size: 13px;
  color: #64748b;
  margin: 0 0 10px 0;
}
.travel-sum{
  background: #fff7ed;
  border: 1px solid rgba(251, 146, 60, .25);
  border-radius: 16px;
  padding: 14px 16px;
  display:flex;
  align-items:center;
  justify-content:space-between;
  margin-top: 10px;
}
.travel-sum .label{font-size:18px;color:#111827;font-weight:700;}
.travel-sum .val{font-size:30px;color:#fb6a20;font-weight:900;}
.travel-kpi{
  background:#f8fafc;
  border:1px solid rgba(148,163,184,.25);
  border-radius: 14px;
  padding: 10px 12px;
  margin-top: 8px;
}
.travel-kpi .k{font-size:13px;color:#64748b;margin-bottom:2px;}
.travel-kpi .v{font-size:18px;font-weight:800;color:#111827;}
</style>
        """,
        unsafe_allow_html=True,
    )


def _auto_download_pdf(pdf_bytes: bytes, filename: str) -> None:
    """One-click PDF download via JS."""
    import base64
    if not pdf_bytes:
        return
    b64 = base64.b64encode(pdf_bytes).decode("utf-8")
    components.html(
        f"""
        <script>
        (function() {{
          const b64 = "{b64}";
          const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
          const blob = new Blob([bytes], {{type: 'application/pdf'}});
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = "{filename}";
          document.body.appendChild(a);
          a.click();
          a.remove();
          setTimeout(() => URL.revokeObjectURL(url), 4000);
        }})();
        </script>
        """,
        height=0,
    )

# ----------------------------
# Helpers
# ----------------------------
def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")

def short_text(s: object, n: int = 10) -> str:
    try:
        t = "" if s is None else str(s)
    except Exception:
        t = ""
    t = t.strip()
    if n <= 0: return ""
    return t if len(t) <= n else (t[:n] + "…")

def parse_sheet_id(s: str) -> str:
    s = (s or "").strip()
    if not s: return ""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", s)
    if m: return m.group(1)
    return s

def normalize_apps_script_url(s: str) -> str:
    s = (s or "").strip()
    if not s: return ""
    if s.startswith("http://") or s.startswith("https://"): return s
    if s.startswith("AKfy"): return f"https://script.google.com/macros/s/{s}/exec"
    return s

def _read_json(path: Path) -> dict:
    if not path.exists(): return {}
    try: return json.loads(path.read_text(encoding="utf-8"))
    except Exception: return {}

def _write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

DEFAULT_CONFIG = {
    "backend": "google",
    "google": {
        "spreadsheet_id": "",
        "submit_sheet_name": "DomesticTrip",
        "draft_sheet_name": "DomesticTrip_Draft",
        "apps_script_url": "",
        "api_key": ""
    }
}

def load_config() -> dict:
    cfg = _read_json(CONFIG_PATH)
    if not cfg:
        cfg = DEFAULT_CONFIG
        _write_json(CONFIG_PATH, cfg)
    merged = DEFAULT_CONFIG | cfg
    merged["google"] = (DEFAULT_CONFIG["google"] | cfg.get("google", {}))
    return merged

def save_config(cfg: dict) -> None:
    _write_json(CONFIG_PATH, cfg)


def get_current_user_email() -> str:
    """Streamlit Community Cloud: read the signed-in user's email when available."""
    try:
        u = getattr(st, "experimental_user", None)
        if u and getattr(u, "email", None):
            return str(u.email).strip()
    except Exception:
        pass
    return ""


def ensure_user_profile(cfg: dict) -> dict:
    """Ensure email -> (user_name, employee_no) mapping.

    Stores mapping in Google Sheet 'Users' via Apps Script when configured.
    Fallback to a local JSON file if cloud isn't configured.
    """
    email = get_current_user_email()
    out = {"email": email, "user_name": "", "employee_no": ""}
    g = (cfg or {}).get("google", {})
    script_url = (g.get("apps_script_url") or "").strip()
    sid = (g.get("spreadsheet_id") or "").strip()
    api_key = (g.get("api_key") or "").strip()

    # Try cloud first
    if email and script_url and sid:
        try:
            prof = cloud_get_user_profile(script_url, sid, email, api_key=api_key)
            if prof:
                out["user_name"] = str(prof.get("user_name") or "").strip()
                out["employee_no"] = str(prof.get("employee_no") or "").strip()
        except Exception:
            pass

    # Local fallback
    local_path = Path(DATA_DIR) / "users.json"
    local = _read_json(local_path) or {}
    if email and email in local:
        out["user_name"] = out["user_name"] or str(local[email].get("user_name") or "").strip()
        out["employee_no"] = out["employee_no"] or str(local[email].get("employee_no") or "").strip()

    # If missing, ask once and persist
    if email and (not out["user_name"] or not out["employee_no"]):
        with st.expander("👤 首次使用：請確認個人資料（將自動帶入表單）", expanded=True):
            nm = st.text_input("使用者姓名", value=out["user_name"], key="travel_profile_name")
            en = st.text_input("員工編號", value=out["employee_no"], key="travel_profile_emp")
            if st.button("儲存個人資料", key="travel_profile_save"):
                out["user_name"] = str(nm).strip()
                out["employee_no"] = str(en).strip()
                if email:
                    local[email] = {"user_name": out["user_name"], "employee_no": out["employee_no"]}
                    _write_json(local_path, local)
                if email and script_url and sid:
                    try:
                        cloud_upsert_user_profile(
                            script_url,
                            sid,
                            {"email": email, "user_name": out["user_name"], "employee_no": out["employee_no"]},
                            api_key=api_key,
                        )
                    except Exception:
                        pass
                st.success("已儲存。之後新增表單會自動帶入。")
                st.rerun()

    return out

def ensure_record_defaults(rec: dict) -> dict:
    defaults = {
        "id": "",
        "status": "draft",
        "filler_name": "",
        "user_email": "",
        "form_date": date.today().isoformat(),
        
        "traveler_name": "",
        "employee_no": "",
        "plan_code": "",
        "purpose_desc": "",
        "travel_route": "",
        "start_time": _now_iso(),
        "end_time": _now_iso(),
        "travel_days": "1",
        
        "is_gov_car": False, "gov_car_no": "",
        "is_taxi": False,
        "is_private_car": False, "private_car_km": 0.0, "private_car_no": "",
        "is_dispatch_car": False,
        "is_hsr": False,
        "is_airplane": False,
        "is_other_transport": False, "other_transport_desc": "",
        
        "estimated_cost": 0.0,
        
        "expense_rows": "[]", # list of dicts
        
        "total_amount": 0.0,
        
        "handler_name": "",
        "project_manager_name": "",
        "dept_manager_name": "",
        "accountant_name": "",
        
        "attachments": "[]",
        "created_at": "",
        "updated_at": "",
        "submitted_at": "",
    }
    out = defaults | (rec or {})
    # Default boolean logic for storage (excel doesn't strictly type bool)
    for k in ["is_gov_car", "is_taxi", "is_private_car", "is_dispatch_car", "is_hsr", "is_airplane", "is_other_transport"]:
        out[k] = bool(out.get(k))
        
    if "attachments" not in out or not out["attachments"]: out["attachments"] = "[]"
    if "expense_rows" not in out or not out["expense_rows"]: out["expense_rows"] = "[]"
    return out

def to_float(x, default=0.0) -> float:
    try:
        if x is None: return float(default)
        import math
        if isinstance(x, float) and math.isnan(x): return float(default)
        if isinstance(x, (int, float)): return float(x)
        s = str(x).strip().replace(",", "").replace("$", "")
        if not s: return float(default)
        return float(s)
    except Exception: return float(default)

def generate_new_id(df: pd.DataFrame, form_date: str) -> str:
    d = form_date.replace("-", "")
    prefix = f"T{d}"
    max_seq = 0
    if df is not None and not df.empty and "id" in df.columns:
        for rid in df["id"].astype(str).tolist():
            if rid.startswith(prefix) and len(rid) >= 12:
                tail = rid[len(prefix):len(prefix)+3]
                if tail.isdigit():
                    max_seq = max(max_seq, int(tail))
    return f"{prefix}{max_seq+1:03d}"

def get_local_df() -> pd.DataFrame:
    conf = load_config()
    draft_sh = conf.get("google", {}).get("draft_sheet_name", "DomesticTrip_Draft")
    submit_sh = conf.get("google", {}).get("submit_sheet_name", "DomesticTrip")
    from storage_excel_travel import cleanup_old_sheets
    cleanup_old_sheets(str(LOCAL_XLSX))
    df = load_all_travel(str(LOCAL_XLSX), draft_sh, submit_sh)
    if df is None or df.empty:
        return pd.DataFrame()

    # 去重：同一 id 若同時存在 draft/submit，以 updated_at 較新者為準
    try:
        df['_updated_at_dt'] = pd.to_datetime(df.get('updated_at', ''), errors='coerce')
        df = df.sort_values(['_updated_at_dt','id'], ascending=[False, False], kind='mergesort')
        df = df.drop_duplicates(subset=['id'], keep='first')
        df = df.sort_values('id', ascending=False, kind='mergesort').reset_index(drop=True)
        df = df.drop(columns=['_updated_at_dt'], errors='ignore')
    except Exception:
        pass

    return df

def get_record_by_id(df: pd.DataFrame, rid: str) -> Optional[dict]:
    if df is None or df.empty: return None
    m = df[df["id"].astype(str) == str(rid)]
    if m.empty: return None
    return ensure_record_defaults(m.iloc[0].to_dict())

def upsert_local_record(rec: dict) -> None:
    rec = ensure_record_defaults(rec)
    rec["updated_at"] = _now_iso()
    if not rec.get("created_at"): rec["created_at"] = rec["updated_at"]

    # Convert true/false booleans to string 'True' or '' for Excel/Cloud consistency
    bool_fields = ["is_gov_car", "is_taxi", "is_private_car", "is_dispatch_car", "is_hsr", "is_airplane", "is_other_transport"]
    for bf in bool_fields:
        val = rec.get(bf)
        if isinstance(val, bool):
            rec[bf] = "True" if val else ""
        elif str(val).lower() == "false":
            rec[bf] = ""
            
    conf = load_config()
    sheet_name = conf.get("google", {}).get("submit_sheet_name", "DomesticTrip") if rec.get("status") == "submitted" else conf.get("google", {}).get("draft_sheet_name", "DomesticTrip_Draft")
    upsert_travel_record(str(LOCAL_XLSX), rec, sheet_name)

def cloud_enabled(cfg: dict) -> bool:
    return cfg.get("backend") == "google" and bool(cfg.get("google", {}).get("spreadsheet_id"))

def cloud_config(cfg: dict) -> dict:
    g = cfg.get("google", {})
    return {
        "spreadsheet_id": parse_sheet_id(g.get("spreadsheet_id", "")),
        "submit_sheet_name": str(g.get("submit_sheet_name", "DomesticTrip")),
        "draft_sheet_name": str(g.get("draft_sheet_name", "DomesticTrip_Draft")),
        "apps_script_url": normalize_apps_script_url(g.get("apps_script_url", "")),
        "api_key": str(g.get("api_key", "")),
    }

def safe_cloud_upsert(cfg: dict, sheet_name: str, rec: dict) -> Tuple[bool, str]:
    if not cloud_enabled(cfg): return True, "disabled"
    g = cloud_config(cfg)
    try:
        cloud_upsert(script_url=g["apps_script_url"], spreadsheet_id=g["spreadsheet_id"], sheet_name=sheet_name, payload=rec, api_key=g["api_key"])
        return True, "OK"
    except Exception as e: return False, str(e)

def safe_cloud_delete(cfg: dict, sheet_name: str, record_id: str) -> Tuple[bool, str]:
    if not cloud_enabled(cfg): return True, "disabled"
    g = cloud_config(cfg)
    try:
        deleted = cloud_delete(script_url=g["apps_script_url"], spreadsheet_id=g["spreadsheet_id"], sheet_name=sheet_name, record_id=record_id, api_key=g["api_key"])
        return True, "deleted" if deleted else "not_found"
    except Exception as e: return False, str(e)

def download_local_excel() -> Tuple[bytes, str]:
    if not LOCAL_XLSX.exists(): _ = get_local_df()
    return LOCAL_XLSX.read_bytes(), LOCAL_XLSX.name

def save_uploaded_files(record_id: str, files: List) -> List[str]:
    folder = ATTACH_DIR / record_id
    folder.mkdir(parents=True, exist_ok=True)
    rel_paths = []
    for f in files:
        name = re.sub(r"[^\w\-.()\[\] ]+", "_", f.name).strip() or "upload"
        target = folder / name
        with open(target, "wb") as out: out.write(f.getbuffer())
        rel_paths.append(str(target.relative_to(APP_DIR)))
    return rel_paths

def resolve_attachment_paths(rel_paths: List[str]) -> List[str]:
    out = []
    for rp in rel_paths:
        try:
            p = (APP_DIR / rp).resolve()
            if p.exists(): out.append(str(p))
        except: pass
    return out

def parse_attachments(cell_json: str) -> List[str]:
    if not cell_json: return []
    try:
        ret = json.loads(cell_json)
        return ret if isinstance(ret, list) else []
    except:
        return []

# ----------------------------
# UI
# ----------------------------

# When user enters this system (or switches from another system), default to NEW form.
# This matches the primary user intent: open system to create a new form.
if st.session_state.get("active_system") != "travel":
    st.session_state["active_system"] = "travel"
    st.session_state["travel_page"] = "new"
    # bump nonce so a new draft is created once per entry
    st.session_state["travel_new_nonce"] = st.session_state.get("travel_new_nonce", 0) + 1
st.markdown("<style>.stButton>button{width:100%;} .tight-buttons .stButton>button{padding-top:0.4rem;padding-bottom:0.4rem;}</style>", unsafe_allow_html=True)

# UI toggles
HIDE_APPROVAL_FIELDS = True  # Hide approval/sign-off fields in forms


def sidebar_settings(cfg: dict) -> dict:
    """Sidebar: keep it simple (no Google Sheet config UI), match the requested style/order."""
    st.sidebar.markdown("## 📌 出差報帳系統")
    st.sidebar.caption("（資料存放於本機 data/，可同步寫入雲端 Google Sheet）")

    g = cloud_config(cfg)

    # --- 工作區（順序：新增 → 草稿 → 已送出） ---
    st.sidebar.markdown("### 📂 工作區")
    current_page = st.session_state.get("travel_page", "new")

    def _go(page: str):
        st.session_state["travel_page"] = page
        if page == "new":
            st.session_state["travel_new_nonce"] = st.session_state.get("travel_new_nonce", 0) + 1
        st.rerun()

    if st.sidebar.button(("✅ " if current_page == "new" else "") + "📝 新增表單（預設）", use_container_width=True, key="t_nav_new"):
        _go("new")
    if st.sidebar.button(("✅ " if current_page == "drafts" else "") + "📄 草稿列表", use_container_width=True, key="t_nav_drafts"):
        _go("drafts")
    if st.sidebar.button(("✅ " if current_page == "list" else "") + "📤 已送出表單列表", use_container_width=True, key="t_nav_submitted"):
        _go("list")

    st.sidebar.divider()

    # --- 資料管理（順序：雲端 → 下載） ---
    st.sidebar.markdown("### 🔗 資料管理")
    open_url = f"https://docs.google.com/spreadsheets/d/{g.get('spreadsheet_id','')}/edit#gid=0" if g.get("spreadsheet_id") else ""
    if open_url:
        st.sidebar.link_button("☁️ 雲端 Excel", open_url, use_container_width=True)
    else:
        st.sidebar.button("☁️ 雲端 Excel", disabled=True, use_container_width=True)

    xls_bytes, xls_name = download_local_excel()
    st.sidebar.download_button(
        "📥 下載 Excel",
        data=xls_bytes,
        file_name=xls_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    return cfg


cfg = sidebar_settings(load_config())


def render_records_table(view: pd.DataFrame, *, scope: str, mode: str):
    """
    將「快速操作」整合進列表：每列右側直接提供 編輯 / 下載PDF / 送出 / 作廢(或刪除)。
    mode:
      - "submitted": 送出 disabled；作廢僅 submitted 可按
      - "drafts": 送出僅 draft 可按；刪除僅 draft/deleted 可按（會刪除草稿）
    """
    if view is None or view.empty:
        st.info("目前篩選後無資料。")
        return

    page_size = st.selectbox("每頁筆數", options=[20, 50, 100], index=0, key=f"{scope}_page_size")
    total_pages = max(1, (len(view) + page_size - 1) // page_size)
    page_no = st.number_input("頁碼", min_value=1, max_value=total_pages, value=1, step=1, key=f"{scope}_page_no")
    start = (page_no - 1) * page_size
    end = start + page_size
    page_df = view.iloc[start:end].copy()

    show = page_df.copy()
    show["狀態標籤"] = show.get("status", "").astype(str).map(
        {"draft": "⚪ draft", "submitted": "🟢 submitted", "void": "🔴 void", "deleted": "⚫ deleted"}
    ).fillna(show.get("status", "").astype(str))
    show["_總金額"] = pd.to_numeric(show.get("total_amount", 0), errors="coerce").fillna(0).apply(lambda v: f"{v:,.0f}")
    show["事由"] = show.get("purpose_desc", "").astype(str).apply(lambda x: short_text(x, 10))

    cols = st.columns([1.2, 1.1, 1.0, 1.0, 1.0, 1.0, 1.3, 1.2, 2.1])
    headers = ["表單ID", "狀態", "日期", "出差人", "計畫編號", "總金額", "事由", "更新時間", "操作"]
    for c, h in zip(cols, headers):
        c.markdown(f"**{h}**")
    st.markdown("<div style='height:0.25rem'></div>", unsafe_allow_html=True)

    g = cloud_config(cfg)

    for _, r in show.iterrows():
        rid = str(r.get("id", ""))
        status = str(r.get("status", ""))
        row = st.columns([1.2, 1.1, 1.0, 1.0, 1.0, 1.0, 1.3, 1.2, 2.1])

        row[0].write(rid)
        row[1].write(str(r.get("狀態標籤", "")))
        row[2].write(str(r.get("form_date", "")))
        row[3].write(str(r.get("traveler_name", "")))
        row[4].write(str(r.get("plan_code", "")))
        row[5].write(str(r.get("_總金額", "")))
        row[6].write(str(r.get("事由", "")))
        row[7].write(str(r.get("updated_at", "")))

        with row[8]:
            b1, b2, b3, b4 = st.columns(4)

            with b1:
                if st.button("編輯", key=f"{scope}_edit_{rid}", use_container_width=True):
                    st.session_state["t_current_id"] = rid
                    st.session_state["travel_page"] = "edit"
                    st.rerun()

            with b2:
                if st.button("下載", key=f"{scope}_dl_{rid}", use_container_width=True):
                    st.session_state["t_current_id"] = rid
                    st.session_state["travel_page"] = "view"
                    st.session_state["auto_download_pdf"] = True
                    st.rerun()

            can_submit = (mode == "drafts") and (status == "draft")
            with b3:
                if st.button("送出", key=f"{scope}_submit_{rid}", disabled=not can_submit, use_container_width=True):
                    rec = get_record_by_id(get_local_df(), rid)
                    if rec:
                        rec["status"] = "submitted"
                        rec["submitted_at"] = _now_iso()
                        upsert_local_record(rec)
                        if cloud_enabled(cfg):
                            safe_cloud_upsert(cfg, g["submit_sheet_name"], rec)
                            safe_cloud_delete(cfg, g["draft_sheet_name"], rid)
                        st.success(f"已送出 {rid}")
                        st.rerun()

            if mode == "submitted":
                label = "作廢"
                can_last = (status == "submitted")
            else:
                label = "刪除"
                can_last = (status in ("draft", "deleted"))

            with b4:
                if st.button(label, key=f"{scope}_last_{rid}", disabled=not can_last, use_container_width=True):
                    if mode == "submitted":
                        rec = get_record_by_id(get_local_df(), rid)
                        if rec:
                            rec["status"] = "void"
                            upsert_local_record(rec)
                            if cloud_enabled(cfg):
                                safe_cloud_upsert(cfg, g["submit_sheet_name"], rec)
                        st.success(f"已作廢 {rid}")
                        st.rerun()
                    else:
                        # 草稿：實際刪除（沿用原本行為）
                        delete_travel_record(str(LOCAL_XLSX), rid, g["draft_sheet_name"])
                        if cloud_enabled(cfg):
                            safe_cloud_delete(cfg, g["draft_sheet_name"], rid)
                        st.success(f"已刪除 {rid}")
                        st.rerun()

# ----------------------------
# Pages
# ----------------------------
def page_list(is_draft_mode=False):
    st.header("草稿列表" if is_draft_mode else "表單列表/查詢")
    inject_travel_ui_css()
    df = get_local_df()
    if df.empty:
        st.info("目前尚無資料。請到左側選『新增出差報帳』開始。")
        return
    df = df.copy()
    for col in ["id", "status", "form_date", "traveler_name", "plan_code", "total_amount", "purpose_desc", "updated_at", "expense_rows"]:
        if col not in df.columns: df[col] = ""
    for col in ["total_amount", "estimated_cost"]:
        if col in df.columns: df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ----- 篩選（含年月區間 / 重設）-----
    # default YM range from data
    base = df.copy()
    if "form_date" not in base.columns:
        base["form_date"] = ""
    _min_ym = str(base["form_date"].astype(str).min())[:7]
    _max_ym = str(base["form_date"].astype(str).max())[:7]
    if not re.match(r"^\d{4}-\d{2}$", _min_ym or ""):
        _min_ym = date.today().strftime("%Y-%m")
    if not re.match(r"^\d{4}-\d{2}$", _max_ym or ""):
        _max_ym = date.today().strftime("%Y-%m")

    st.session_state.setdefault("travel_filter_status", "draft" if is_draft_mode else "(全部)")
    st.session_state.setdefault("travel_filter_traveler", "")
    st.session_state.setdefault("travel_filter_plan", "")
    st.session_state.setdefault("travel_filter_id", "")
    st.session_state.setdefault("travel_filter_start", _min_ym)
    st.session_state.setdefault("travel_filter_end", _max_ym)

    if st.button("重設篩選", key="travel_reset_filters"):
        st.session_state["travel_filter_status"] = "draft" if is_draft_mode else "(全部)"
        st.session_state["travel_filter_traveler"] = ""
        st.session_state["travel_filter_plan"] = ""
        st.session_state["travel_filter_id"] = ""
        st.session_state["travel_filter_start"] = _min_ym
        st.session_state["travel_filter_end"] = _max_ym
        st.rerun()

    c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.2, 1.2])
    if is_draft_mode:
        status = c1.selectbox("狀態", options=["draft", "deleted"], key="travel_filter_status")
    else:
        status = c1.selectbox("狀態", options=["(全部)", "submitted", "void"], key="travel_filter_status")
    traveler = c2.text_input("出差人包含", key="travel_filter_traveler")
    plan = c3.text_input("計畫編號包含", key="travel_filter_plan")
    rid_kw = c4.text_input("表單ID", key="travel_filter_id")

    c5, c6 = st.columns([1.2, 1.2])
    start_ym = c5.text_input("起始年月(YYYY-MM)", key="travel_filter_start")
    end_ym = c6.text_input("結束年月(YYYY-MM)", key="travel_filter_end")

    if not re.match(r"^\d{4}-\d{2}$", str(start_ym).strip()):
        start_ym = st.session_state["travel_filter_start"] = _min_ym
    if not re.match(r"^\d{4}-\d{2}$", str(end_ym).strip()):
        end_ym = st.session_state["travel_filter_end"] = _max_ym

    view = df.copy()
    if is_draft_mode:
        # Pre-filter for drafts
        view = view[view["status"].astype(str).isin(["draft", "deleted"])]
    else:
        # Pre-filter for submitted/void
        view = view[view["status"].astype(str).isin(["submitted", "void"])]

    if status != "(全部)":
        view = view[view["status"].astype(str) == status]
    if traveler:
        view = view[view.get("traveler_name","").astype(str).str.contains(traveler, na=False)]
    if plan:
        view = view[view.get("plan_code","").astype(str).str.contains(plan, na=False)]
    if rid_kw:
        view = view[view.get("id","").astype(str).str.contains(rid_kw, na=False)]

    # YM range filter
    def _in_ym(d):
        d = str(d)
        return len(d) >= 7 and start_ym <= d[:7] <= end_ym
    if "form_date" in view.columns:
        view = view[view["form_date"].astype(str).apply(_in_ym)]

    view = view.sort_values(by=["form_date","id"], ascending=[False, False])
    vshow = view.copy()
    vshow["狀態標籤"] = vshow.get("status","").astype(str).map({"draft":"⚪ draft","submitted":"🟢 submitted","void":"🔴 void","deleted":"⚫ deleted"}).fillna(vshow.get("status","").astype(str))
        # 表單列表（可直接在每列右側操作：編輯/下載/送出/作廢或刪除）
    render_records_table(
        view,
        scope="travel_drafts" if is_draft_mode else "travel_list",
        mode="drafts" if is_draft_mode else "submitted",
    )

    
    # Totals footer（KPI 條：交通/膳雜/住宿/其它 + 預估 + 筆數；並顯示報支總計）
    total_all = pd.to_numeric(view.get("total_amount", 0), errors="coerce").fillna(0).sum()
    total_est = pd.to_numeric(view.get("estimated_cost", 0), errors="coerce").fillna(0).sum()

    total_transport = total_per_diem = total_accom = total_other = 0.0
    if "expense_rows" in view.columns:
        for cell in view["expense_rows"].tolist():
            try:
                if cell is None or (isinstance(cell, float) and pd.isna(cell)):
                    rows = []
                else:
                    rows = json.loads(cell or "[]")
                if not isinstance(rows, list):
                    rows = []
            except Exception:
                rows = []
            for r in rows:
                if not isinstance(r, dict):
                    continue
                total_transport += to_float(r.get("transport_amt", 0.0))
                total_per_diem += to_float(r.get("per_diem_amt", 0.0))
                total_accom += to_float(r.get("accommodation_amt", 0.0))
                total_other += to_float(r.get("other_amt", 0.0))

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.markdown(f"<div class='travel-kpi'><div class='k'>交通費合計</div><div class='v'>NT$ {total_transport:,.0f}</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='travel-kpi'><div class='k'>膳雜費合計</div><div class='v'>NT$ {total_per_diem:,.0f}</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='travel-kpi'><div class='k'>住宿費合計</div><div class='v'>NT$ {total_accom:,.0f}</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='travel-kpi'><div class='k'>其它支出合計</div><div class='v'>NT$ {total_other:,.0f}</div></div>", unsafe_allow_html=True)
    k5.markdown(f"<div class='travel-kpi'><div class='k'>出差費預估合計</div><div class='v'>NT$ {total_est:,.0f}</div></div>", unsafe_allow_html=True)
    k6.markdown(f"<div class='travel-kpi'><div class='k'>筆數</div><div class='v'>{len(view):,}</div></div>", unsafe_allow_html=True)

    st.markdown(
        f"<div class='travel-sum'><div class='label'>報支總計：</div><div class='val'>NT$ {total_all:,.0f}</div></div>",
        unsafe_allow_html=True,
    )

def page_new():
    """Auto-create a new travel draft and jump to edit (no extra click)."""
    nonce = st.session_state.get("travel_new_nonce", 0)
    created_nonce = st.session_state.get("travel_new_created_nonce", -1)
    if created_nonce == nonce and st.session_state.get("t_current_id"):
        st.session_state["travel_page"] = "edit"
        st.rerun()

    cfg = load_config()
    prof = ensure_user_profile(cfg)
    df = get_local_df()
    form_date = date.today().isoformat()
    rid = generate_new_id(df, form_date)
    rec = ensure_record_defaults({"id": rid, "form_date": form_date, "status": "draft"})
    if prof.get("email"):
        rec["user_email"] = prof.get("email")
    # default-fill fields
    if prof.get("user_name"):
        rec["traveler_name"] = prof.get("user_name")
        rec["filler_name"] = prof.get("user_name")
    if prof.get("employee_no"):
        rec["employee_no"] = prof.get("employee_no")
    rec["created_at"] = _now_iso()
    rec["updated_at"] = rec["created_at"]
    upsert_local_record(rec)

    st.session_state["t_current_id"] = rid
    st.session_state["travel_new_created_nonce"] = nonce
    st.session_state["travel_page"] = "edit"
    st.rerun()

def page_drafts():
    page_list(is_draft_mode=True)


def page_edit():
    rid = st.session_state.get("t_current_id", "")
    cfg = load_config()
    df = get_local_df()
    rec = get_record_by_id(df, rid)
    if not rec:
        st.error("找不到此表單。")
        return

    inject_travel_ui_css()

    # Helper for session state mapping
    def ss_bind(k, val=None):
        sk = f"trec_{rid}_{k}"
        if sk not in st.session_state:
            st.session_state[sk] = rec.get(k, val) if val is not None else rec.get(k)
        return sk

    # Title (match the softer style)
    st.markdown(
        f"""
        <div style="display:flex;align-items:center;gap:12px;margin-top:4px;">
          <div style="font-size:44px;">🧑‍💼</div>
          <div>
            <div class="travel-title">新增出差申請與報支單</div>
            <div class="travel-sub">帶有 <b>*</b> 號的欄位為必填項目。請依序完成下方卡片內容。</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ----------------------------
    # 1) 基本資料與行程事由
    # ----------------------------
    with st.container(border=True):
        st.markdown("### 👤 1. 基本資料與行程事由")
        st.caption("＊為必填；簽核欄位已隱藏（將於印出 PDF 後由主管簽署）。")

        # show signed-in email when available
        _email = get_current_user_email()
        if _email:
            st.markdown(f"<div class='travel-sub'>登入信箱：<b>{_email}</b></div>", unsafe_allow_html=True)
            st.session_state[ss_bind("user_email", _email)] = _email

        c1, c2, c3 = st.columns([1.2, 1.0, 1.0])
        with c1:
            st.text_input("出差人姓名 *", value=st.session_state[ss_bind("traveler_name", "")], key=ss_bind("traveler_name"))
        with c2:
            st.text_input("員工編號 *", value=st.session_state[ss_bind("employee_no", "")], key=ss_bind("employee_no"))
        with c3:
            st.text_input("計畫編號", value=st.session_state[ss_bind("plan_code", "")], key=ss_bind("plan_code"))

        c4, c5 = st.columns([2.2, 1.4])
        with c4:
            st.text_area("出差事由 / 工作摘要 *", value=st.session_state[ss_bind("purpose_desc", "")], key=ss_bind("purpose_desc"), height=68)
        with c5:
            st.text_input("出差起訖地點 *", value=st.session_state[ss_bind("travel_route", "")], key=ss_bind("travel_route"))

    # ----------------------------
    # 2) 出差期間 + 出差費預估
    # ----------------------------
    with st.container(border=True):
        st.markdown("### ⏰ 2. 出差期間")
        sc1, sc2, sc3, sc4, sc5 = st.columns([1.1, 1.0, 1.1, 1.0, 1.0])

        try:
            s_dt = datetime.fromisoformat(st.session_state[ss_bind("start_time", _now_iso())])
        except Exception:
            s_dt = datetime.now()
        try:
            e_dt = datetime.fromisoformat(st.session_state[ss_bind("end_time", _now_iso())])
        except Exception:
            e_dt = datetime.now()

        sd = sc1.date_input("起始日期", value=s_dt.date(), key=f"tui_{rid}_sd")
        st_t = sc2.time_input("起始時間", value=s_dt.time().replace(second=0, microsecond=0), key=f"tui_{rid}_st")
        ed = sc3.date_input("結束日期", value=e_dt.date(), key=f"tui_{rid}_ed")
        et_t = sc4.time_input("結束時間", value=e_dt.time().replace(second=0, microsecond=0), key=f"tui_{rid}_et")

        st.session_state[ss_bind("start_time")] = datetime.combine(sd, st_t).isoformat(timespec="seconds")
        st.session_state[ss_bind("end_time")] = datetime.combine(ed, et_t).isoformat(timespec="seconds")
        # travel days (整數)
        try:
            days_val = int(float(str(st.session_state[ss_bind("travel_days", "1")] or 1).strip() or 1))
        except Exception:
            days_val = 1
        days_val = int(sc5.number_input("總出差天數", min_value=0, value=int(days_val), step=1, key=f"tui_{rid}_days"))
        st.session_state[ss_bind("travel_days")] = str(days_val)

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        st.markdown("### 💰 出差費預估")
        # IMPORTANT: do NOT assign to st.session_state for a widget key after instantiation.
        # Ensure default is set before creating the widget, then just read its value later.
        _est_key = ss_bind("estimated_cost", 0)

        def _coerce_int(v, default=0):
            try:
                s = str(v).strip()
                if s == "" or s.lower() in {"nan", "none"}:
                    return int(default)
                return int(float(s))
            except Exception:
                return int(default)

        # Streamlit will prefer st.session_state[_est_key] over the provided `value=`.
        # If the key ever contains a non-numeric (e.g. ''), number_input will crash when comparing to min_value.
        # So we *always* normalize the key to an int BEFORE instantiating the widget.
        st.session_state[_est_key] = _coerce_int(
            st.session_state.get(_est_key, rec.get("estimated_cost", 0)),
            default=0,
        )

        st.number_input(
            "出差費預估（元）",
            min_value=0,
			value=int(st.session_state.get(_est_key, 0) or 0),
			step=1,
            format="%d",
            key=_est_key,
        )

    # ----------------------------
    # 3) 交通工具與里程
    # ----------------------------
    with st.container(border=True):
        st.markdown("### 🚗 3. 交通工具與里程")
        st.caption("可多選；若有公務車/私車公用/其他，請補充車號或說明。")

        options = ["公務車", "計程車", "私車公用", "派車服務", "高鐵", "國內飛機", "其他"]
        # derive default selection from bool flags
        default_sel = []
        if st.session_state[ss_bind("is_gov_car", False)]: default_sel.append("公務車")
        if st.session_state[ss_bind("is_taxi", False)]: default_sel.append("計程車")
        if st.session_state[ss_bind("is_private_car", False)]: default_sel.append("私車公用")
        if st.session_state[ss_bind("is_dispatch_car", False)]: default_sel.append("派車服務")
        if st.session_state[ss_bind("is_hsr", False)]: default_sel.append("高鐵")
        if st.session_state[ss_bind("is_airplane", False)]: default_sel.append("國內飛機")
        if st.session_state[ss_bind("is_other_transport", False)]: default_sel.append("其他")

        sel = st.multiselect("此次行程中使用的交通工具（可多選）", options=options, default=default_sel, key=f"tui_{rid}_transport_sel")

        # map to bool flags
        st.session_state[ss_bind("is_gov_car")] = ("公務車" in sel)
        st.session_state[ss_bind("is_taxi")] = ("計程車" in sel)
        st.session_state[ss_bind("is_private_car")] = ("私車公用" in sel)
        st.session_state[ss_bind("is_dispatch_car")] = ("派車服務" in sel)
        st.session_state[ss_bind("is_hsr")] = ("高鐵" in sel)
        st.session_state[ss_bind("is_airplane")] = ("國內飛機" in sel)
        st.session_state[ss_bind("is_other_transport")] = ("其他" in sel)

        a, b = st.columns(2)
        with a:
            if "公務車" in sel:
                st.text_input("車號（公務車）", value=st.session_state[ss_bind("gov_car_no", "")], key=ss_bind("gov_car_no"))
        with b:
            if "私車公用" in sel:
                st.number_input("里程數（公里）", min_value=0.0, value=float(st.session_state[ss_bind("private_car_km", 0.0)] or 0.0), step=1.0, key=ss_bind("private_car_km"))
                st.text_input("車號（私車）", value=st.session_state[ss_bind("private_car_no", "")], key=ss_bind("private_car_no"))
            if "其他" in sel:
                st.text_input("其他說明", value=st.session_state[ss_bind("other_transport_desc", "")], key=ss_bind("other_transport_desc"))

    # ----- Init editor df in session_state -----
    editor_df_key = f"trec_{rid}_expense_editor_df"
    if editor_df_key not in st.session_state:
        try:
            rows = json.loads(st.session_state[ss_bind("expense_rows", "[]")] or "[]")
            if not isinstance(rows, list):
                rows = []
        except Exception:
            rows = []

        def _parse_date(v: str):
            try:
                s = (v or "").strip()
                if not s:
                    return None
                if re.match(r"^\d{1,2}/\d{1,2}$", s):
                    y = date.today().year
                    m, d = s.split("/")
                    return date(y, int(m), int(d))
                return datetime.fromisoformat(s).date()
            except Exception:
                return None

        editor_rows = []
        for r in rows:
            editor_rows.append({
                "日期": _parse_date(r.get("date_md") or ""),
                "起訖地點": (r.get("route") or ""),
                "車別": (r.get("transport_type") or ""),
                "交通費": int(float(str(r.get("transport_amt") or 0).replace(',', '') or 0)),
                "膳雜費": int(float(str(r.get("per_diem_amt") or 0).replace(',', '') or 0)),
                "住宿費": int(float(str(r.get("accommodation_amt") or 0).replace(',', '') or 0)),
                "其它支出": int(float(str(r.get("other_amt") or 0).replace(',', '') or 0)),
                "單據編號": (r.get("receipt_no") or ""),
            })
        if not editor_rows:
            editor_rows = [{"日期": None, "起訖地點": "", "車別": "", "交通費": 0, "膳雜費": 0, "住宿費": 0, "其它支出": 0, "單據編號": ""}]
        st.session_state[editor_df_key] = pd.DataFrame(editor_rows)

    vehicle_opts = ["", "計程車", "高鐵", "公務車", "私車公用", "派車服務", "國內飛機", "其他"]

    def _to_int(v) -> int:
        try:
            if v is None:
                return 0
            if isinstance(v, (int,)):
                return int(v)
            s = str(v).strip().replace(',', '')
            if s == "":
                return 0
            return int(round(float(s)))
        except Exception:
            return 0

    def _sync_session_to_rec(edited_df: pd.DataFrame, uploaded_files):
        # sync simple fields
        for k in list(rec.keys()):
            sk = f"trec_{rid}_{k}"
            if sk in st.session_state:
                rec[k] = st.session_state[sk]

        # always keep user_email
        if get_current_user_email():
            rec["user_email"] = get_current_user_email()

        # sync details from editor
        new_rows = []
        total_transport = total_per_diem = total_accom = total_other = 0
        # NOTE: Do NOT call edited_df.fillna("") here.
        # Some numeric columns use pandas nullable integers (Int64) and cannot be filled with "".
        # Instead, handle NA values cell-by-cell.
        import pandas as _pd

        for _, rr in edited_df.iterrows():
            dd = rr.get("日期")
            date_md = ""
            if dd is None:
                date_md = ""
            else:
                try:
                    if _pd.isna(dd) and not isinstance(dd, (str,)):
                        date_md = ""
                    elif isinstance(dd, datetime):
                        date_md = dd.date().isoformat()
                    elif isinstance(dd, date):
                        date_md = dd.isoformat()
                    else:
                        # could be string like '2026-03-05 00:00:00' or '2026-03-05T00:00:00'
                        date_md = str(dd).strip().split("T")[0].split(" ")[0]
                except Exception:
                    date_md = str(dd).strip().split("T")[0].split(" ")[0]

            def _s(v) -> str:
                if v is None:
                    return ""
                try:
                    if _pd.isna(v):
                        return ""
                except Exception:
                    pass
                s = str(v)
                return "" if s == "<NA>" else s.strip()

            route = _s(rr.get("起訖地點", ""))
            veh = _s(rr.get("車別", ""))
            rec_no = _s(rr.get("單據編號", ""))
            t_amt = _to_int(rr.get("交通費"))
            p_amt = _to_int(rr.get("膳雜費"))
            a_amt = _to_int(rr.get("住宿費"))
            o_amt = _to_int(rr.get("其它支出"))
            if not (date_md or route or veh or rec_no or t_amt or p_amt or a_amt or o_amt):
                continue
            new_rows.append({
                "date_md": date_md,
                "route": route,
                "transport_type": veh,
                "transport_amt": t_amt,
                "per_diem_amt": p_amt,
                "accommodation_amt": a_amt,
                "other_amt": o_amt,
                "receipt_no": rec_no,
            })
            total_transport += t_amt
            total_per_diem += p_amt
            total_accom += a_amt
            total_other += o_amt

        pending = json.dumps(new_rows, ensure_ascii=False)
        rec["expense_rows"] = pending
        st.session_state[ss_bind("expense_rows")] = pending
        st.session_state[ss_bind("_pending_expense_rows")] = pending

        total_all = int(total_transport + total_per_diem + total_accom + total_other)
        rec["total_amount"] = total_all
        st.session_state[ss_bind("total_amount")] = total_all

        # sync attachments
        if ss_bind("attachments") not in st.session_state:
            st.session_state[ss_bind("attachments")] = rec.get("attachments", "[]")
        existing_att = parse_attachments(st.session_state[ss_bind("attachments")])
        if uploaded_files:
            new_rel = save_uploaded_files(rid, uploaded_files)
            merged_dict = {x: True for x in existing_att}
            for p in new_rel:
                merged_dict[p] = True
            existing_att = list(merged_dict.keys())
            st.session_state[ss_bind("attachments")] = json.dumps(existing_att, ensure_ascii=False)
        rec["attachments"] = st.session_state[ss_bind("attachments")]

        rec["updated_at"] = _now_iso()
        return total_transport, total_per_diem, total_accom, total_other, total_all

    # ---- form: no rerun while editing table ----
    with st.form(key=f"travel_detail_form_{rid}", clear_on_submit=False):
        with st.container(border=True):
            st.markdown("### 🧾 4. 差旅費報支單據明細")
            st.caption("明細表格採『提交制』：填寫完成後，按下下方按鈕（儲存草稿/送出/下載PDF）才會一次性更新與保存。")
            # Streamlit's data_editor is strict about dataframe dtypes when column_config is supplied.
            # Normalize dtypes to avoid StreamlitAPIException (type compatibility).
            df_render = st.session_state[editor_df_key].copy()
            # Ensure expected columns exist
            for col in ["日期", "起訖地點", "車別", "交通費", "膳雜費", "住宿費", "其它支出", "單據編號"]:
                if col not in df_render.columns:
                    df_render[col] = "" if col in ["起訖地點", "車別", "單據編號"] else 0

            # 日期 -> datetime64 (enables calendar picker reliably)
            df_render["日期"] = pd.to_datetime(df_render["日期"], errors="coerce")
            # Selectbox: NaN -> "" so value is always in options
            df_render["車別"] = df_render["車別"].fillna("").astype(str)
            # Numeric columns -> Int64
            for ncol in ["交通費", "膳雜費", "住宿費", "其它支出"]:
                df_render[ncol] = pd.to_numeric(df_render[ncol], errors="coerce").fillna(0).round(0).astype("Int64")

            edited = st.data_editor(
                df_render,
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "日期": st.column_config.DateColumn(format="YYYY-MM-DD"),
                    "起訖地點": st.column_config.TextColumn(),
                    "車別": st.column_config.SelectboxColumn(options=vehicle_opts),
                    "交通費": st.column_config.NumberColumn(min_value=0, step=1, format="%d"),
                    "膳雜費": st.column_config.NumberColumn(min_value=0, step=1, format="%d"),
                    "住宿費": st.column_config.NumberColumn(min_value=0, step=1, format="%d"),
                    "其它支出": st.column_config.NumberColumn(min_value=0, step=1, format="%d"),
                    "單據編號": st.column_config.TextColumn(),
                },
                key=f"tui_{rid}_editor",
            )

        # Persist latest table edits so they are used when saving/submitting/downloading PDF.
        st.session_state[editor_df_key] = edited

        with st.container(border=True):
            st.markdown("### 📎 附件（收據/發票/照片/PDF）")
            existing_att = parse_attachments(st.session_state.get(ss_bind("attachments"), rec.get("attachments", "[]")))
            if existing_att:
                existing_abs = resolve_attachment_paths(existing_att)
                for p in existing_abs:
                    st.write(f"- {Path(p).name}")
            uploaded_files = st.file_uploader(
                "新增附件（可多選）",
                type=["pdf", "png", "jpg", "jpeg", "webp"],
                accept_multiple_files=True,
                key=f"t_uploader_{rid}",
            )

        st.divider()
        send_pdf = st.checkbox("送出後寄送 PDF 到我的信箱", value=False, key=f"tui_{rid}_send_pdf")
        a1, a2, a3, a4 = st.columns([1.0, 1.2, 1.0, 1.0])
        save_clicked = a1.form_submit_button("儲存草稿", use_container_width=True)
        submit_clicked = a2.form_submit_button("確認無誤並送出", type="primary", use_container_width=True)
        pdf_clicked = a3.form_submit_button("下載PDF", use_container_width=True)
        back_clicked = a4.form_submit_button("返回列表", use_container_width=True)

    # after submit: persist edited df
    if save_clicked or submit_clicked or pdf_clicked or back_clicked:
        st.session_state[editor_df_key] = edited

    if back_clicked:
        st.session_state["travel_page"] = "list"
        st.rerun()

    if save_clicked or submit_clicked or pdf_clicked:
        t1, t2, t3, t4, total_all = _sync_session_to_rec(edited, uploaded_files)

        # show totals (KPI bar)
        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(f"<div class='travel-kpi'><div class='k'>交通費合計</div><div class='v'>NT$ {t1:,.0f}</div></div>", unsafe_allow_html=True)
        k2.markdown(f"<div class='travel-kpi'><div class='k'>膳雜費合計</div><div class='v'>NT$ {t2:,.0f}</div></div>", unsafe_allow_html=True)
        k3.markdown(f"<div class='travel-kpi'><div class='k'>住宿費合計</div><div class='v'>NT$ {t3:,.0f}</div></div>", unsafe_allow_html=True)
        k4.markdown(f"<div class='travel-kpi'><div class='k'>其它支出合計</div><div class='v'>NT$ {t4:,.0f}</div></div>", unsafe_allow_html=True)
        st.markdown(f"<div class='travel-sum'><div class='label'>報支總計：</div><div class='val'>NT$ {total_all:,.0f}</div></div>", unsafe_allow_html=True)

        # Save local first
        if submit_clicked:
            missing = []
            if not str(rec.get("traveler_name", "")).strip():
                missing.append("出差人姓名")
            if not str(rec.get("employee_no", "")).strip():
                missing.append("員工編號")
            if not str(rec.get("purpose_desc", "")).strip():
                missing.append("出差事由/工作摘要")
            if not str(rec.get("travel_route", "")).strip():
                missing.append("出差起訖地點")
            if missing:
                st.error("以下必填欄位尚未填寫：" + "、".join(missing))
                return
            rec["status"] = "submitted"
            rec["submitted_at"] = _now_iso()
        else:
            if rec.get("status") != "submitted":
                rec["status"] = "draft"

        upsert_local_record(rec)

        # cloud
        if cloud_enabled(cfg):
            g = cloud_config(cfg)
            if submit_clicked:
                safe_cloud_upsert(cfg, g["submit_sheet_name"], rec)
                safe_cloud_delete(cfg, g["draft_sheet_name"], rid)
            else:
                safe_cloud_upsert(cfg, g["draft_sheet_name"], rec)

        if submit_clicked:
            # optional email PDF
            if send_pdf:
                try:
                    email = get_current_user_email() or str(rec.get("user_email") or "").strip()
                    if email and cloud_enabled(cfg):
                        g = cloud_config(cfg)
                        paths = resolve_attachment_paths(parse_attachments(rec.get("attachments", "[]")))
                        pdf_bytes = build_pdf_bytes(rec, attachment_paths=paths)
                        import base64
                        cloud_send_pdf_email(
                            g["apps_script_url"],
                            g["spreadsheet_id"],
                            to_email=email,
                            subject=f"出差報帳表單 {rid}",
                            filename=f"出差報帳_{rid}.pdf",
                            pdf_base64=base64.b64encode(pdf_bytes).decode("ascii"),
                            body_text="已為您附上出差報帳 PDF。",
                            api_key=g.get("api_key", ""),
                        )
                        st.info("已寄送 PDF 到您的信箱。")
                except Exception as e:
                    st.warning(f"寄信失敗（不影響送出）：{e}")

            st.success("表單已送出！")
            st.session_state["travel_page"] = "list"
            st.rerun()

        if pdf_clicked:
            try:
                paths = resolve_attachment_paths(parse_attachments(rec.get("attachments", "[]")))
                pdf_bytes = build_pdf_bytes(rec, attachment_paths=paths)
                _auto_download_pdf(pdf_bytes, f"出差報帳_{rid}.pdf")
                st.download_button(
                    "若未自動下載，請點此下載 PDF",
                    data=pdf_bytes,
                    file_name=f"出差報帳_{rid}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    key=f"t_fallback_{rid}",
                )
                st.success("已儲存草稿並準備下載 PDF。")
            except Exception as e:
                st.error(f"PDF 產生失敗：{e}")
        else:
            st.success("已儲存草稿！")

def page_view():
    rid = st.session_state.get("t_current_id", "")
    st.header(f"檢視/下載出差表單：{rid}")
    df = get_local_df()
    rec = get_record_by_id(df, rid)
    if not rec:
        st.error("找不到此表單。")
        return
        
    st.json(rec, expanded=False)
    
    st.info("處理 PDF 中...")
    paths = resolve_attachment_paths(parse_attachments(rec.get("attachments", "[]")))
    pdf_bytes = build_pdf_bytes(rec, attachment_paths=paths)
    if pdf_bytes:
        filename = f"TravelVoucher_{rid}.pdf"
        st.download_button(
            label="下載出差單 PDF",
            data=pdf_bytes,
            file_name=filename,
            mime="application/pdf",
            use_container_width=True,
            type="primary"
        )
    else:
        st.error("無法產生 PDF。")
        
    if st.button("返回編輯", use_container_width=True):
        st.session_state["travel_page"] = "edit"
        st.rerun()

current_page = st.session_state.get("travel_page", "new")
if current_page == "list": page_list(is_draft_mode=False)
elif current_page == "new": page_new()
elif current_page == "drafts": page_drafts()
elif current_page == "edit": page_edit()
elif current_page == "view": page_view()
