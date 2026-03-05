import sys
from datetime import datetime
from storage_apps_script import list_records, upsert_record

if len(sys.argv) < 4:
    print("Usage: python cloud_smoketest.py <script_url> <spreadsheet_id> <sheet_name> [api_key]")
    sys.exit(1)

script_url = sys.argv[1]
spreadsheet_id = sys.argv[2]
sheet_name = sys.argv[3]
api_key = sys.argv[4] if len(sys.argv) >= 5 else ""

print("[1] list_records test")
rows = list_records(script_url, spreadsheet_id, sheet_name, api_key)
print(f"OK list: {len(rows)} rows")

print("[2] upsert_record test")
test_id = f"SMOKE-{datetime.now().strftime('%Y%m%d-%H%M%S')}"
upsert_record(script_url, spreadsheet_id, sheet_name, {
    "id": test_id,
    "status": "draft",
    "form_date": datetime.now().strftime("%Y-%m-%d"),
    "filler_name": "smoke_test",
    "purpose_desc": "cloud test"
}, api_key)
print("OK upsert:", test_id)
