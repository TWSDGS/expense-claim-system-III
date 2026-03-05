1) 請先把 storage_apps_script.py 覆蓋到專案根目錄。
2) 請把 run_windows.bat 覆蓋後重啟（固定使用 8501）。
3) 側欄雲端設定：
   - backend=google
   - spreadsheet_id=1i8Iw8dTfrKGpCOdxMXl5d2QMgOD7VbA84UEPRjBc_zw
   - submit_sheet_name=vouchers
   - draft_sheet_name=vouchers（或在 Google Sheet 新增「草稿列表」）
   - apps_script_url=你的 Web App /exec URL
   - api_key=若 Code.gs 有設 API_KEY 才要填
4) 在專案資料夾執行：
   .venv\Scripts\python.exe cloud_smoketest.py "<apps_script_url>" "1i8Iw8dTfrKGpCOdxMXl5d2QMgOD7VbA84UEPRjBc_zw" "vouchers"

如果失敗，錯誤訊息會直接指出是：權限、URL、sheet 名稱、或網路問題。
