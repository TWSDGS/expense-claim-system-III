# storage_apps_script.py
# Google Apps Script Web App helper for reading/writing Google Sheets

from __future__ import annotations

import json
from typing import Any, Dict, Optional

import requests


class AppsScriptError(RuntimeError):
    pass


def call_script(
    script_url: str,
    body: Dict[str, Any],
    timeout: int = 30,
) -> Dict[str, Any]:
    """POST JSON to Apps Script Web App and return parsed JSON."""
    if not script_url:
        raise AppsScriptError("Apps Script URL is empty")
    resp = requests.post(
        script_url,
        data=json.dumps(body),
        headers={"Content-Type": "application/json"},
        timeout=timeout,
    )
    resp.raise_for_status()
    data = resp.json()
    if not isinstance(data, dict):
        raise AppsScriptError(f"Unexpected response: {data!r}")
    if not data.get("ok", False):
        raise AppsScriptError(str(data.get("error", "Unknown error")))
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
    return None


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
