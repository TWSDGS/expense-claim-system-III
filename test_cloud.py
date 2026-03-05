import json, pathlib, requests

cfg = json.loads(pathlib.Path("data/config.json").read_text(encoding="utf-8"))
print("backend =", cfg.get("backend"))

g = cfg.get("google", {})
body = {
    "action": "list",
    "spreadsheetId": g.get("spreadsheet_id", ""),
    "sheetName": g.get("submit_sheet_name", "申請表單"),
}
if g.get("api_key"):
    body["apiKey"] = g["api_key"]

r = requests.post(g.get("apps_script_url", ""), json=body, timeout=20)
print("HTTP =", r.status_code)
print("Content-Type =", r.headers.get("content-type"))
print("Body(head) =", r.text[:800])
