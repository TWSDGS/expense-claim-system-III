/**
 * 支出報帳系統 - Google Sheets 寫入/讀取（Apps Script Web App）
 *
 * 部署：
 * 1) script.google.com 新建專案，貼上此 Code.gs
 * 2) (可選) 設定 API Key：專案設定 → 指令碼屬性
 *    KEY: API_KEY   VALUE: 你自訂的一串字
 * 3) 部署 → 新增部署 → 網頁應用程式
 *    - 以誰執行：我
 *    - 誰可存取：任何知道連結的人（或你的網域內）
 * 4) 複製 Web App URL，貼到 Streamlit 設定的 Apps Script URL
 *
 * 這支 Web App 透過 POST JSON 提供：
 * - action=list   讀取指定工作表所有資料（自第3列起）
 * - action=upsert 依 payload.id 寫入（若已存在則更新，否則新增）
 * - action=delete 依 id 刪除一筆
 *
 * Sheet 欄位：
 * - 第1列：欄位 key（英文，對應 Python payload keys）
 * - 第2列：欄位名稱（中文，方便人工檢視，可自行修改）
 */

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _getApiKey() {
  return PropertiesService.getScriptProperties().getProperty("API_KEY") || "";
}

function _requireKey(body) {
  var required = _getApiKey();
  if (!required) return true; // 沒設定就不驗證
  var provided = (body && body.apiKey) ? String(body.apiKey) : "";
  if (provided !== required) {
    throw new Error("apiKey mismatch");
  }
  return true;
}

function _headers() {
  return [
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
    "submitted_at"
  ];
}

function _headersZh() {
  return [
    "表單編號",
    "狀態",
    "填表人",
    "填表日期",
    "計畫編號",
    "用途說明",
    "付款模式",
    "受款類型",
    "員工姓名",
    "員工編號",
    "廠商名稱",
    "廠商地址",
    "廠商受款人",
    "借支沖銷",
    "借支金額",
    "沖銷金額",
    "結餘繳回",
    "補領金額",
    "發票/收據號碼",
    "未稅金額",
    "稅額",
    "總計",
    "承辦人",
    "計畫主持人",
    "部門主管",
    "會計",
    "附件",
    "建立時間",
    "更新時間",
    "送出時間"
  ];
}

// Users sheet (for Streamlit Community Cloud email -> profile mapping)
function _userHeaders() {
  return ["email", "user_name", "employee_no", "created_at", "updated_at"]; 
}

function _ensureUsersSheet(ss) {
  var sh = ss.getSheetByName("Users");
  if (!sh) sh = ss.insertSheet("Users");
  // init headers if empty
  if (sh.getLastRow() < 1 || sh.getLastColumn() === 0) {
    sh.clear();
    sh.getRange(1, 1, 1, _userHeaders().length).setValues([_userHeaders()]);
  } else {
    var existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    while (existing.length && String(existing[existing.length-1]).trim() === "") existing.pop();
    var merged = existing.slice();
    var headers = _userHeaders();
    for (var i=0; i<headers.length; i++) {
      if (merged.indexOf(headers[i]) === -1) merged.push(headers[i]);
    }
    if (merged.length !== existing.length) {
      sh.getRange(1, 1, 1, merged.length).setValues([merged]);
    }
  }
  return sh;
}

function _findUserRowByEmail(sh, email) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var emailCol = headers.indexOf("email") + 1;
  if (emailCol <= 0) return -1;
  var vals = sh.getRange(2, emailCol, last-1, 1).getValues();
  for (var i=0; i<vals.length; i++) {
    if (String(vals[i][0]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      return i + 2;
    }
  }
  return -1;
}

function _userProfileFromRow(headers, row) {
  var out = {};
  for (var i=0; i<headers.length; i++) {
    out[headers[i]] = (row && row.length > i) ? row[i] : "";
  }
  return out;
}

function _userRowFromProfile(headers, profile) {
  var row = [];
  for (var i=0; i<headers.length; i++) {
    var k = headers[i];
    row.push((profile && profile.hasOwnProperty(k)) ? profile[k] : "");
  }
  return row;
}

function _ensureSheet(ss, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  return sh;
}

function _ensureHeaders(sh, extraKeys) {
  var headers = _headers();
  // 允許 payload 帶來新欄位，自動擴充 header（避免欄位錯位）
  if (extraKeys && extraKeys.length) {
    for (var i=0; i<extraKeys.length; i++) {
      var k = String(extraKeys[i]);
      if (k && headers.indexOf(k) === -1) headers.push(k);
    }
  }

  var lastCol = sh.getLastColumn();
  var needInit = (sh.getLastRow() < 2 || lastCol === 0);

  if (needInit) {
    sh.clear();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    var zh = _headersZh();
    // 若中文欄位數不足，補空白
    while (zh.length < headers.length) zh.push("");
    sh.getRange(2, 1, 1, headers.length).setValues([zh.slice(0, headers.length)]);
    return headers;
  }

  // 讀取現有第1列 headers
  var existing = sh.getRange(1, 1, 1, Math.max(1, lastCol)).getValues()[0];
  // 去除尾端空白
  while (existing.length && String(existing[existing.length-1]).trim() === "") existing.pop();

  // 若現有 header 與我們的不一致，採用「以現有為主 + 補齊缺少欄位」
  // 避免覆蓋人工改過的欄位順序
  var merged = existing.slice();
  for (var j=0; j<headers.length; j++) {
    if (merged.indexOf(headers[j]) === -1) merged.push(headers[j]);
  }

  // 若 merged 有擴充，寫回第1列；第2列不足則補空白
  if (merged.length !== existing.length) {
    sh.getRange(1, 1, 1, merged.length).setValues([merged]);
    var zh2 = sh.getRange(2, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    while (zh2.length < merged.length) zh2.push("");
    sh.getRange(2, 1, 1, merged.length).setValues([zh2.slice(0, merged.length)]);
  }

  return merged;
}

function _payloadFromRow(headers, row) {
  var out = {};
  for (var i=0; i<headers.length; i++) {
    var k = headers[i];
    out[k] = (row && row.length > i) ? row[i] : "";
  }
  return out;
}

function _rowFromPayload(headers, payload) {
  var row = [];
  for (var i=0; i<headers.length; i++) {
    var k = headers[i];
    row.push((payload && payload.hasOwnProperty(k)) ? payload[k] : "");
  }
  return row;
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    _requireKey(body);

    var action = String(body.action || "");
    var spreadsheetId = String(body.spreadsheetId || "");
    var sheetName = String(body.sheetName || "草稿列表");

    if (!spreadsheetId) throw new Error("spreadsheetId is required");

    var ss = SpreadsheetApp.openById(spreadsheetId);

    // ---- Users mapping ----
    if (action === "user_get") {
      var email = String(body.email || "").trim();
      if (!email) throw new Error("email is required");
      var ush = _ensureUsersSheet(ss);
      var headersU = ush.getRange(1, 1, 1, ush.getLastColumn()).getValues()[0];
      while (headersU.length && String(headersU[headersU.length-1]).trim() === "") headersU.pop();
      var rowIdx = _findUserRowByEmail(ush, email);
      if (rowIdx < 0) return _json({ok:true, profile: null});
      var row = ush.getRange(rowIdx, 1, 1, headersU.length).getValues()[0];
      return _json({ok:true, profile: _userProfileFromRow(headersU, row)});
    }

    if (action === "user_upsert") {
      var profile = body.profile || {};
      var email2 = String(profile.email || "").trim();
      if (!email2) throw new Error("profile.email is required");
      var ush2 = _ensureUsersSheet(ss);
      var headersU2 = ush2.getRange(1, 1, 1, ush2.getLastColumn()).getValues()[0];
      while (headersU2.length && String(headersU2[headersU2.length-1]).trim() === "") headersU2.pop();
      profile.updated_at = new Date().toISOString();
      if (!profile.created_at) profile.created_at = profile.updated_at;
      var rowIdx2 = _findUserRowByEmail(ush2, email2);
      if (rowIdx2 < 0) {
        ush2.appendRow(_userRowFromProfile(headersU2, profile));
        return _json({ok:true, upserted:true, mode:"insert"});
      }
      ush2.getRange(rowIdx2, 1, 1, headersU2.length).setValues([_userRowFromProfile(headersU2, profile)]);
      return _json({ok:true, upserted:true, mode:"update"});
    }

    if (action === "send_email") {
      var to = String(body.to || "").trim();
      var subject = String(body.subject || "報帳表單 PDF");
      var filename = String(body.filename || "report.pdf");
      var pdf64 = String(body.pdf_base64 || "");
      var bodyText = String(body.body || "");
      if (!to) throw new Error("to is required");
      if (!pdf64) throw new Error("pdf_base64 is required");
      var bytes = Utilities.base64Decode(pdf64);
      var blob = Utilities.newBlob(bytes, "application/pdf", filename);
      MailApp.sendEmail({
        to: to,
        subject: subject,
        body: bodyText || "附件為 PDF 報表。",
        attachments: [blob],
      });
      return _json({ok:true, sent:true});
    }

    var sh = _ensureSheet(ss, sheetName);

    if (action === "list") {
      var vals = sh.getDataRange().getValues();
      if (vals.length <= 2) return _json({ok:true, rows: []}); // Row1 headers, Row2 zh
      var headers = vals[0];
      // 去掉尾端空欄
      while (headers.length && String(headers[headers.length-1]).trim() === "") headers.pop();
      var rows = [];
      for (var r=2; r<vals.length; r++) {
        rows.push(_payloadFromRow(headers, vals[r]));
      }
      return _json({ok:true, rows: rows});
    }

    if (action === "upsert") {
      var payload = body.payload || {};
      var id = String(payload.id || "");
      if (!id) throw new Error("payload.id is required");

      // 確保 headers（並依 payload keys 擴充）
      var keys = Object.keys(payload);
      var headers2 = _ensureHeaders(sh, keys);

      payload.updated_at = new Date().toISOString();
      if (!payload.created_at) payload.created_at = payload.updated_at;

      var last = sh.getLastRow();
      if (last < 3) {
        sh.insertRowBefore(3);
        sh.getRange(3, 1, 1, headers2.length).setValues([_rowFromPayload(headers2, payload)]);
        return _json({ok:true, upserted:true, mode:"insert_top"});
      }

      var idRange = sh.getRange(3, 1, last-2, 1).getValues();
      for (var i=0; i<idRange.length; i++) {
        if (String(idRange[i][0]) === id) {
          sh.getRange(i + 3, 1, 1, headers2.length).setValues([_rowFromPayload(headers2, payload)]);
          return _json({ok:true, upserted:true, mode:"update"});
        }
      }

      sh.insertRowBefore(3);
      sh.getRange(3, 1, 1, headers2.length).setValues([_rowFromPayload(headers2, payload)]);
      return _json({ok:true, upserted:true, mode:"insert_top"});
    }

    if (action === "delete") {
      var idToDelete = String(body.id || "");
      if (!idToDelete) throw new Error("id is required for delete");

      var last2 = sh.getLastRow();
      if (last2 < 3) return _json({ok:true, deleted:false});

      var idRange2 = sh.getRange(3, 1, last2-2, 1).getValues();
      for (var j=0; j<idRange2.length; j++) {
        if (String(idRange2[j][0]) === idToDelete) {
          sh.deleteRow(j + 3);
          return _json({ok:true, deleted:true});
        }
      }
      return _json({ok:true, deleted:false});
    }

    throw new Error("Unknown action: " + action);
  } catch (err) {
    return _json({ok:false, error: String(err)});
  }
}

/**
 * 草稿 7 天未送出提醒（需手動在 Apps Script 內建立時間觸發器，或呼叫 setupDraftReminderTrigger）
 * 規則：掃描所有工作表中 status==draft 或工作表名稱包含「草稿」，若 updated_at <= today-7 且 user_email 有值則寄提醒。
 */
function setupDraftReminderTrigger() {
  // 每天 09:00 執行一次
  var triggers = ScriptApp.getProjectTriggers();
  for (var i=0; i<triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runDraftReminder') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('runDraftReminder').timeBased().everyDays(1).atHour(9).create();
}

function runDraftReminder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var cutoff = new Date(now.getTime() - 7*24*60*60*1000);
  var sheets = ss.getSheets();

  for (var s=0; s<sheets.length; s++) {
    var sh = sheets[s];
    var name = sh.getName();
    var vals = sh.getDataRange().getValues();
    if (vals.length <= 2) continue;
    var headers = vals[0];
    while (headers.length && String(headers[headers.length-1]).trim() === '') headers.pop();
    var idxStatus = headers.indexOf('status');
    var idxUpdated = headers.indexOf('updated_at');
    var idxEmail = headers.indexOf('user_email');
    var idxId = headers.indexOf('id');

    // 只處理有必要欄位的 sheet
    if (idxUpdated < 0 || idxEmail < 0) continue;

    for (var r=2; r<vals.length; r++) {
      var row = vals[r];
      var status = (idxStatus >= 0) ? String(row[idxStatus] || '') : '';
      var userEmail = String(row[idxEmail] || '').trim();
      if (!userEmail) continue;
      if (status && status !== 'draft') continue;
      if (!status && name.indexOf('草稿') === -1 && name.toLowerCase().indexOf('draft') === -1) continue;
      var updatedStr = String(row[idxUpdated] || '').trim();
      if (!updatedStr) continue;
      var updated = new Date(updatedStr);
      if (isNaN(updated.getTime())) continue;
      if (updated > cutoff) continue;

      var formId = (idxId >= 0) ? String(row[idxId] || '') : '';
      var subject = '提醒：草稿超過 7 天未送出' + (formId ? ('（' + formId + '）') : '');
      var body = '您有一筆草稿已超過 7 天未送出或刪除。\n\n' +
                 '工作表：' + name + '\n' +
                 (formId ? ('表單編號：' + formId + '\n') : '') +
                 '最後更新：' + updatedStr + '\n\n' +
                 '請回到系統確認是否需要送出或刪除。';
      MailApp.sendEmail(userEmail, subject, body);
    }
  }
}
