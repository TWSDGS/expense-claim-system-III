# storage_apps_script.py
# Google Apps Script Web App helper for reading/writing Google Sheets

from __future__ import annotations

import json
from typing import Any, Dict

import requests


class AppsScriptError(RuntimeError):
    pass


def _preview(text: str, limit: int = 300) -> str:
    text = (text or "").strip().replace("\n", " ")
    return text[:limit]


def call_script(
    script_url: str,
    body: Dict[str, Any],
    timeout: int = 30,
) -> Dict[str, Any]:
    """POST JSON to Apps Script Web App and return parsed JSON.

    若 Apps Script 部署權限或 URL 有問題，會拋出帶有診斷資訊的 AppsScriptError。
    """
    if not script_url:
        raise AppsScriptError("Apps Script URL is empty")

    try:
        resp = requests.post(
            script_url,
            data=json.dumps(body, ensure_ascii=False),
            headers={"Content-Type": "application/json"},
            timeout=timeout,
            allow_redirects=True,
        )
    except requests.RequestException as e:
        raise AppsScriptError(f"Network error: {type(e).__name__}: {e}") from e

    status = resp.status_code
    content_type = (resp.headers.get("Content-Type") or "").lower()

    if status >= 400:
        raise AppsScriptError(
            f"HTTP {status}; body={_preview(resp.text)}"
        )

    try:
        data = resp.json()
    except ValueError:
        hint = ""
        if "text/html" in content_type:
            hint = " (回傳 HTML，通常是 Apps Script 權限或 URL 不是 /exec)"
        raise AppsScriptError(
            f"Non-JSON response{hint}; Content-Type={content_type}; body={_preview(resp.text)}"
        )

    if not isinstance(data, dict):
        raise AppsScriptError(f"Unexpected response type: {type(data).__name__}, data={data!r}")

    if not data.get("ok", False):
        err = data.get("error", "Unknown error")
        raise AppsScriptError(f"Apps Script error: {err}")

    return data


def list_records(
    script_url: str,
    spreadsheet_id: str,
    sheet_name: str,
    api_key: str = "",
) -> list[dict]:
    body = {
        "action": "list",
        "spreadsheetId": spreadsheet_id,
        "sheetName": sheet_name,
    }
    if api_key:
        body["apiKey"] = api_key
    data = call_script(script_url, body)
    return list(data.get("rows", []))


def upsert_record(
    script_url: str,
    spreadsheet_id: str,
    sheet_name: str,
    payload: dict,
    api_key: str = "",
) -> None:
    body = {
        "action": "upsert",
        "spreadsheetId": spreadsheet_id,
        "sheetName": sheet_name,
        "payload": payload,
    }
    if api_key:
        body["apiKey"] = api_key
    call_script(script_url, body)


def delete_record(
    script_url: str,
    spreadsheet_id: str,
    sheet_name: str,
    record_id: str,
    api_key: str = "",
) -> bool:
    body = {
        "action": "delete",
        "spreadsheetId": spreadsheet_id,
        "sheetName": sheet_name,
        "id": record_id,
    }
    if api_key:
        body["apiKey"] = api_key
    data = call_script(script_url, body)
    return bool(data.get("deleted", False))


def get_user_profile(
    script_url: str,
    spreadsheet_id: str,
    email: str,
    api_key: str = "",
) -> dict | None:
    """Fetch user profile by email from a 'Users' sheet.

    Apps Script must implement action='user_get'.
    """
    body = {
        "action": "user_get",
        "spreadsheetId": spreadsheet_id,
        "email": email,
    }
    if api_key:
        body["apiKey"] = api_key
    data = call_script(script_url, body)
    return data.get("profile")


def upsert_user_profile(
    script_url: str,
    spreadsheet_id: str,
    profile: dict,
    api_key: str = "",
) -> None:
    """Upsert user profile to 'Users' sheet.

    Apps Script must implement action='user_upsert'.
    """
    body = {
        "action": "user_upsert",
        "spreadsheetId": spreadsheet_id,
        "profile": profile,
    }
    if api_key:
        body["apiKey"] = api_key
    call_script(script_url, body)


def send_pdf_email(
    script_url: str,
    spreadsheet_id: str,
    to_email: str,
    subject: str,
    filename: str,
    pdf_base64: str,
    body_text: str = "",
    api_key: str = "",
) -> None:
    """Send a PDF to email via Apps Script.

    Apps Script must implement action='send_email'.
    """
    body = {
        "action": "send_email",
        "spreadsheetId": spreadsheet_id,
        "to": to_email,
        "subject": subject,
        "filename": filename,
        "pdf_base64": pdf_base64,
        "body": body_text or "",
    }
    if api_key:
        body["apiKey"] = api_key
    call_script(script_url, body, timeout=60)
