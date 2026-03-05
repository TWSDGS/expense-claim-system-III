/**
 * 出差報帳系統 - Google Sheets 寫入/讀取（Apps Script Web App）
 *
 * 部署：
 * 1) script.google.com 新建專案，貼上此 Code_travel.gs
 * 2) (可選) 設定 API Key：專案設定 → 指令碼屬性
 *    KEY: API_KEY   VALUE: 你自訂的一串字
 * 3) 部署 → 新增部署 → 網頁應用程式
 *    - 以誰執行：我
 *    - 誰可存取：任何知道連結的人（或你的網域內）
 * 4) 複製 Web App URL，貼到 Streamlit 設定的 Apps Script URL
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
    throw new Error("Unauthorized: API key mismatch");
  }
  return true;
}

function _headers() {
  return [
    "id", 
    "status", 
    "filler_name", 
    "form_date", 
    "traveler_name", 
    "plan_code", 
    "purpose_desc", 
    "travel_route",
    "start_time", 
    "end_time", 
    "travel_days",
    "is_gov_car", "gov_car_no", 
    "is_taxi", 
    "is_private_car", "private_car_km", "private_car_no",
    "is_dispatch_car", 
    "is_hsr", 
    "is_airplane", 
    "is_other_transport", "other_transport_desc",
    "estimated_cost", 
    "expense_rows", 
    "total_amount",
    "handler_name", "project_manager_name", "dept_manager_name", "accountant_name",
    "attachments", 
    "created_at", 
    "updated_at", 
    "submitted_at"
  ];
}

function _headersZh() {
  return [
    "表單編號", "狀態", "填表人", "填表日期", "出差人", "計畫編號", "出差事由", "出差行程",
    "出差起始時間", "出差結束時間", "出差天數",
    "公務車", "公務車號", "計程車",
    "私車", "私車公里數", "私車車號",
    "派車", "高鐵", "飛機", "其他交通", "其他交通說明",
    "預估總花費", "出差明細(JSON)", "總金額",
    "經手人", "計畫主持人", "部門主管", "會計",
    "附件", "建立時間", "更新時間", "送出時間"
  ];
}

function _ensureSheet(ss, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  var headers = _headers();
  var headersZh = _headersZh();
  var lastRow = sh.getLastRow();
  
  if (lastRow === 0) {
    sh.appendRow(headers);
    sh.appendRow(headersZh);
  } else {
    // Forcefully sync English headers to row 1 to prevent column shifts
    var currentHeaders = sh.getRange(1, 1, 1, sh.getLastColumn() || 1).getValues()[0];
    var headerStr1 = currentHeaders.slice(0, 5).join(",");
    var targetStr1 = headers.slice(0, 5).join(",");
    
    if (headerStr1 !== targetStr1) {
      // Overwrite the first two rows with correct schema
      if (lastRow === 1) {
        sh.appendRow(headersZh);
      }
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.getRange(2, 1, 1, headersZh.length).setValues([headersZh]);
    }
  }
  return sh;
}

function _rowFromPayload(p) {
  var h = _headers();
  var row = [];
  for (var i=0; i<h.length; i++) {
    var k = h[i];
    row.push((p && (k in p)) ? p[k] : "");
  }
  return row;
}

function _payloadFromRow(values) {
  var h = _headers();
  var obj = {};
  for (var i=0; i<h.length; i++) {
    obj[h[i]] = (i < values.length) ? values[i] : "";
  }
  return obj;
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
    var sheetName = String(body.sheetName || "DomesticTrip_Draft");

    if (!spreadsheetId) throw new Error("spreadsheetId is required");

    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sh = _ensureSheet(ss, sheetName);

    if (action === "list") {
      var vals = sh.getDataRange().getValues();
      if (vals.length <= 2) return _json({ok:true, rows: []}); // Row 1: EN, Row 2: ZH
      var rows = [];
      for (var r=2; r<vals.length; r++) {
        rows.push(_payloadFromRow(vals[r]));
      }
      return _json({ok:true, rows: rows});
    }

    if (action === "upsert") {
      var payload = body.payload || {};
      var id = String(payload.id || "");
      if (!id) throw new Error("payload.id is required");

      payload.updated_at = new Date().toISOString();
      if (!payload.created_at) payload.created_at = payload.updated_at;

      var last = sh.getLastRow();
      var idRange = (last >= 3) ? sh.getRange(3, 1, last-2, 1).getValues() : [];
      var targetRow = -1;
      for (var i=0; i<idRange.length; i++) {
        if (String(idRange[i][0]) === id) { targetRow = i + 3; break; }
      }

      var rowValues = _rowFromPayload(payload);
      if (targetRow === -1) {
        sh.appendRow(rowValues);
      } else {
        sh.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
      }

      return _json({ok:true});
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
