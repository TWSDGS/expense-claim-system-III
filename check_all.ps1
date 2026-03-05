param(
    [string]$ScriptUrl = "https://script.google.com/macros/s/AKfycbxK2t639xgj2-jjeYfalrJ8AgTh8v5UayRglu0eKmXB3lOJGuNjaPgAV7DrAjI2Walx/exec",
    [string]$SpreadsheetId = "1i8Iw8dTfrKGpCOdxMXl5d2QMgOD7VbA84UEPRjBc_zw",
    [string]$SheetName = "vouchers",
    [string]$ApiKey = ""
)

$ErrorActionPreference = "Stop"

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("========== " + $Title + " ==========") -ForegroundColor Cyan
}

function Write-Ok {
    param([string]$Msg)
    Write-Host ("[OK] " + $Msg) -ForegroundColor Green
}

function Write-Warn {
    param([string]$Msg)
    Write-Host ("[WARN] " + $Msg) -ForegroundColor Yellow
}

function Write-Err {
    param([string]$Msg)
    Write-Host ("[ERR] " + $Msg) -ForegroundColor Red
}

function Invoke-JsonPost {
    param(
        [string]$Url,
        [hashtable]$Obj
    )
    $body = $Obj | ConvertTo-Json -Depth 10
    Invoke-RestMethod -Method Post -Uri $Url -ContentType "application/json" -Body $body -TimeoutSec 30
}

try {
    # ----------------------------
    # 0) Basic argument validation
    # ----------------------------
    $ScriptUrl = $ScriptUrl.Trim()
    if (-not ($ScriptUrl -match '^https://')) {
        throw "ScriptUrl must start with https:// and end with /exec"
    }
    if (-not ($ScriptUrl -match '/exec$')) {
        Write-Warn "ScriptUrl does not end with /exec. Please confirm this is Web App URL."
    }

    Write-Section "0) Parameters"
    Write-Host ("ScriptUrl     = " + $ScriptUrl)
    Write-Host ("SpreadsheetId = " + $SpreadsheetId)
    Write-Host ("SheetName     = " + $SheetName)
    if ([string]::IsNullOrWhiteSpace($ApiKey)) {
        Write-Host "ApiKey        = (not provided)"
    } else {
        Write-Host "ApiKey        = (provided)"
    }

    # ----------------------------
    # 1) local port 8501
    # ----------------------------
    Write-Section "1) Local Streamlit port 8501"
    $conn = Get-NetTCPConnection -LocalPort 8501 -State Listen -ErrorAction SilentlyContinue
    if ($null -ne $conn) {
        $pidVal = ($conn | Select-Object -First 1).OwningProcess
        Write-Ok ("Port 8501 is LISTENING. PID=" + $pidVal)
    } else {
        Write-Warn "Port 8501 is not listening. Start Streamlit first if needed."
    }

    # ----------------------------
    # 2) localhost health
    # ----------------------------
    Write-Section "2) localhost:8501 health check"
    try {
        $r = Invoke-WebRequest -Uri "http://localhost:8501" -UseBasicParsing -TimeoutSec 5
        Write-Ok ("localhost:8501 reachable. HTTP " + $r.StatusCode)
    } catch {
        Write-Warn ("localhost:8501 not reachable: " + $_.Exception.Message)
    }

    # ----------------------------
    # 3) Apps Script GET /exec
    # ----------------------------
    Write-Section "3) Apps Script GET /exec"
    $ping = Invoke-RestMethod -Method Get -Uri $ScriptUrl -TimeoutSec 30

    if ($ping -is [string]) {
        $trim = $ping.Trim()
        if ($trim.StartsWith("{")) {
            $obj = $trim | ConvertFrom-Json
            if ($null -ne $obj.ok) {
                Write-Ok ("GET /exec JSON ok=" + $obj.ok)
            } else {
                Write-Warn "GET /exec returned JSON string but no 'ok' field."
            }
        } else {
            Write-Err "GET /exec returned non-JSON text (possible permission issue or wrong URL)."
            $previewLen = [Math]::Min(300, $trim.Length)
            Write-Host ($trim.Substring(0, $previewLen))
            throw "Invalid GET /exec response"
        }
    } else {
        if ($null -ne $ping.ok) {
            Write-Ok ("GET /exec JSON ok=" + $ping.ok)
        } else {
            Write-Warn "GET /exec returned object but no 'ok' field."
        }
    }

    # ----------------------------
    # 4) Apps Script POST list
    # ----------------------------
    Write-Section "4) Apps Script POST list"
    $listPayload = @{
        action        = "list"
        spreadsheetId = $SpreadsheetId
        sheetName     = $SheetName
    }
    if (-not [string]::IsNullOrWhiteSpace($ApiKey)) {
        $listPayload.apiKey = $ApiKey
    }

    $listRes = Invoke-JsonPost -Url $ScriptUrl -Obj $listPayload
    if ($listRes.ok -ne $true) {
        $errText = ""
        if ($null -ne $listRes.error) { $errText = [string]$listRes.error }
        Write-Err ("POST list failed. " + $errText)
        throw "POST list failed"
    }

    $count = 0
    if ($null -ne $listRes.rows) {
        $count = @($listRes.rows).Count
    }
    Write-Ok ("POST list success. rows=" + $count)

    # ----------------------------
    # 5) Apps Script POST upsert
    # ----------------------------
    Write-Section "5) Apps Script POST upsert"
    $testId = "SMOKE-" + (Get-Date -Format "yyyyMMdd-HHmmss")
    $utcNow = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    $upsertPayload = @{
        action        = "upsert"
        spreadsheetId = $SpreadsheetId
        sheetName     = $SheetName
        payload       = @{
            id                    = $testId
            status                = "submitted"
            form_date             = (Get-Date -Format "yyyy-MM-dd")
            plan_code             = "SMOKE"
            purpose_desc          = "cloud smoke test"
            payment_mode          = "transfer"
            payee_type            = "employee"
            employee_name         = "smoketest"
            employee_no           = "T000"
            vendor_name           = ""
            vendor_address        = ""
            vendor_payee_name     = ""
            is_advance_offset     = "N"
            advance_amount        = 0
            offset_amount         = 0
            balance_refund_amount = 0
            supplement_amount     = 0
            receipt_no            = ""
            amount_total          = 1
            handler_name          = "smoketest"
            project_manager_name  = ""
            dept_manager_name     = ""
            accountant_name       = ""
            attachments           = ""
            submitted_at          = $utcNow
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($ApiKey)) {
        $upsertPayload.apiKey = $ApiKey
    }

    $upsertRes = Invoke-JsonPost -Url $ScriptUrl -Obj $upsertPayload
    if ($upsertRes.ok -ne $true) {
        $errText2 = ""
        if ($null -ne $upsertRes.error) { $errText2 = [string]$upsertRes.error }
        Write-Err ("POST upsert failed. " + $errText2)
        throw "POST upsert failed"
    }
    Write-Ok ("POST upsert success. test id=" + $testId)

    # ----------------------------
    # 6) Verify by listing again
    # ----------------------------
    Write-Section "6) Verify write by list again"
    $listRes2 = Invoke-JsonPost -Url $ScriptUrl -Obj $listPayload
    if ($listRes2.ok -ne $true) {
        Write-Err "Second list failed."
        throw "Second list failed"
    }

    $rows2 = @()
    if ($null -ne $listRes2.rows) {
        $rows2 = @($listRes2.rows)
    }

    $hit = $rows2 | Where-Object { $_.id -eq $testId }
    if ($null -ne $hit) {
        Write-Ok ("Verification success: found id=" + $testId)
    } else {
        Write-Warn "Upsert returned ok, but id not found in list. Refresh sheet and check manually."
    }

    Write-Section "Done"
    Write-Ok "All checks completed."
}
catch {
    Write-Host ""
    Write-Err ("Check aborted: " + $_.Exception.Message)
    exit 1
}
