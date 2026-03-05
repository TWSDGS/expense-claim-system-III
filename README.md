# 支出報帳系統（Streamlit + 本機 Excel + Google Sheets 雲端 + PDF）

## 特色
- 新增／查詢／編輯／草稿／送出（送出後可下載 PDF）
- 本機備份：`data/vouchers.xlsx`
- 雲端（可選）：Google Sheet + Apps Script Web App
  - **送出工作表名稱**：預設 `申請表單`
  - **草稿工作表名稱**：預設 `草稿列表`
- PDF：以 `templates/voucher_bg.png` 作為底圖，疊上欄位文字與勾選（A4）＋自動合併附件（PDF/圖片）
- 已內建中文字型（避免 PDF 亂碼）

---

## 1) 一鍵執行

### Windows
雙擊 `run_windows.bat`

### Mac / Linux
```bash
chmod +x run_mac_linux.sh
./run_mac_linux.sh
```

---

## 2) 介面操作重點
1. 左側欄「儲存模式」選：
   - **本機 Excel**：只寫入本機備份
   - **Google Sheet（Apps Script）**：本機 + 雲端
2. 左側欄「Google Sheet 設定」貼上：
   - Sheet ID（或整個 Sheet 網址）
   - Apps Script URL（或部署 ID `AKfy...`）
   - 送出工作表名稱：`申請表單`
   - 草稿工作表名稱：`草稿列表`
3. 進入「新增表單」建立後開始填寫：
   - 最上方有「**填表人**」欄位（草稿篩選用）
   - 付款對象三選一：員工姓名／借支充抵／廠商付款（互斥）
4. 產生 PDF：
   - 右側欄按「產生 PDF」→「下載 PDF」
5. 儲存：
   - 「儲存草稿」：寫本機 + 雲端草稿工作表
   - 「送出」：寫本機 + 雲端送出工作表（並嘗試自動刪除雲端草稿）

---

## 3) Apps Script 部署（雲端寫入用）
1. 到 `script.google.com` 新增專案
2. 把 `apps_script/Code.gs` 全部貼上
3. （可選）設定 API Key：專案設定 → 指令碼屬性 `API_KEY`
4. 部署 → 新增部署 → 網頁應用程式
   - 以誰執行：我
   - 誰可存取：任何知道連結的人（或你的網域內）
5. 複製 Web App URL，貼到系統左側欄的 Apps Script URL

> **注意**：如果你更新 `Code.gs`（例如新增欄位），要重新部署新的版本（或更新部署）。

---

## 4) 檔案結構
- `app.py`：主介面
- `pdf_gen.py`：PDF 生成與附件合併
- `storage_excel.py`：本機 Excel 讀寫
- `storage_apps_script.py`：Apps Script Web App 呼叫
- `data/vouchers.xlsx`：本機備份
- `data/attachments/<表單ID>/`：附件存放
