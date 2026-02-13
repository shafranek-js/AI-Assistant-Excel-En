param(
    [string]$InputDir = ".\vba_unpacked_utf8",
    [string]$OutputPath = "",
    [string]$Track = "dev",
    [bool]$UpdateLatestAlias = $true
)

$ErrorActionPreference = "Stop"

function New-BuildFileName {
    param([string]$TrackName)

    $t = $TrackName.ToLowerInvariant()
    if ($t -notmatch "^[a-z0-9_-]+$") {
        throw "Invalid track name '$TrackName'. Allowed: a-z, 0-9, underscore, hyphen."
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmm"
    $datePart = $stamp.Split("_")[0]
    $pattern = "AI_Assistant_${t}_${datePart}_*_b*.xlsm"
    $existing = Get-ChildItem -Path . -Filter $pattern -ErrorAction SilentlyContinue

    $maxBuild = 0
    foreach ($f in $existing) {
        if ($f.Name -match "_b(\d{2})\.xlsm$") {
            $n = [int]$matches[1]
            if ($n -gt $maxBuild) { $maxBuild = $n }
        }
    }

    $nextBuild = "b{0:d2}" -f ($maxBuild + 1)
    "AI_Assistant_${t}_${stamp}_${nextBuild}.xlsm"
}

function Get-VbaBody {
    param([string]$Path)

    $lines = Get-Content -LiteralPath $Path
    $cleanLines = $lines | ForEach-Object { ($_ -replace "^\uFEFF", "") }
    $body = $cleanLines | Where-Object { $_ -notmatch "^\s*Attribute\s+VB_" }
    ($body -join "`r`n").Trim()
}

function Add-Control {
    param(
        $Designer,
        [string]$ProgId,
        [string]$Name,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [hashtable]$Props = @{}
    )

    $ctrl = $Designer.Controls.Add($ProgId)
    $ctrl.Name = $Name
    $ctrl.Left = $Left
    $ctrl.Top = $Top
    $ctrl.Width = $Width
    $ctrl.Height = $Height

    foreach ($k in $Props.Keys) {
        try {
            $ctrl.$k = $Props[$k]
        }
        catch {
            # Some MSForms controls do not expose all properties on every host.
        }
    }

    $ctrl
}

function Add-FormCode {
    param(
        $Component,
        [string]$FormCodePath
    )

    $body = Get-VbaBody -Path $FormCodePath
    if ([string]::IsNullOrWhiteSpace($body)) {
        return
    }
    $Component.CodeModule.AddFromString($body)
}

function Set-FormSize {
    param(
        $Component,
        $Designer,
        [double]$Width,
        [double]$Height
    )

    # Designer size affects what you see in VBE.
    try { $Designer.Width = $Width } catch {}
    try { $Designer.Height = $Height } catch {}

    # Component properties are a fallback for hosts where Designer size is not persisted.
    try { $Component.Properties.Item("Width").Value = $Width } catch {}
    try { $Component.Properties.Item("Height").Value = $Height } catch {}
}

function Add-StandardModule {
    param(
        $VBProject,
        [string]$Name,
        [string]$ModulePath
    )

    $module = $VBProject.VBComponents.Add(1)
    $module.Name = $Name
    $body = Get-VbaBody -Path $ModulePath
    if (-not [string]::IsNullOrWhiteSpace($body)) {
        $module.CodeModule.AddFromString($body)
    }
}

$requiredFiles = @(
    "modAIHelper.bas",
    "modExcelHelper.bas",
    "modMain.bas",
    "frmChat.frm",
    "frmSettings.frm",
    "ThisWorkbook.cls",
    "Sheet1.cls"
)

foreach ($f in $requiredFiles) {
    $p = Join-Path $InputDir $f
    if (-not (Test-Path -LiteralPath $p)) {
        throw "Missing required file: $p"
    }
}

$outputPathResolved = $OutputPath
if ([string]::IsNullOrWhiteSpace($outputPathResolved)) {
    $outputPathResolved = New-BuildFileName -TrackName $Track
}

$outputFull = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $outputPathResolved))
if (Test-Path -LiteralPath $outputFull) {
    Remove-Item -LiteralPath $outputFull -Force
}

$excel = $null
$wb = $null
$latestPath = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    $vb = $wb.VBProject

    Add-StandardModule -VBProject $vb -Name "modAIHelper" -ModulePath (Join-Path $InputDir "modAIHelper.bas")
    Add-StandardModule -VBProject $vb -Name "modExcelHelper" -ModulePath (Join-Path $InputDir "modExcelHelper.bas")
    Add-StandardModule -VBProject $vb -Name "modMain" -ModulePath (Join-Path $InputDir "modMain.bas")

    $frmChat = $vb.VBComponents.Add(3)
    $frmChat.Name = "frmChat"
    $d = $frmChat.Designer
    try { $d.Caption = "AI Assistant for Excel" } catch {}
    Set-FormSize -Component $frmChat -Designer $d -Width 540 -Height 535

    Add-Control $d "Forms.TextBox.1" "txtChat" 12 12 500 315 @{
        MultiLine = $true
        EnterKeyBehavior = $true
        WordWrap = $true
        ScrollBars = 2
    } | Out-Null

    Add-Control $d "Forms.Label.1" "lblStatus" 12 334 500 18 @{
        Caption = "Ready"
    } | Out-Null

    Add-Control $d "Forms.TextBox.1" "txtInput" 12 356 500 58 @{
        MultiLine = $true
        EnterKeyBehavior = $true
        WordWrap = $true
        ScrollBars = 2
    } | Out-Null

    Add-Control $d "Forms.OptionButton.1" "optCloud" 12 434 70 18 @{
        Caption = "Cloud"
        Value = $true
    } | Out-Null

    Add-Control $d "Forms.OptionButton.1" "optLocal" 86 434 70 18 @{
        Caption = "Local"
    } | Out-Null

    Add-Control $d "Forms.ComboBox.1" "cmbModel" 160 432 170 21 @{} | Out-Null

    Add-Control $d "Forms.Label.1" "lblLocalModel" 160 434 170 18 @{
        Caption = "LM Studio (auto)"
        Visible = $false
    } | Out-Null

    Add-Control $d "Forms.CheckBox.1" "chkIncludeData" 334 434 82 18 @{
        Caption = "Include Data"
        Value = $true
    } | Out-Null

    Add-Control $d "Forms.CheckBox.1" "chkPreviewCommands" 420 434 92 18 @{
        Caption = "Preview"
        Value = $false
    } | Out-Null

    Add-Control $d "Forms.CommandButton.1" "btnSend" 12 460 90 28 @{
        Caption = "Send"
    } | Out-Null

    Add-Control $d "Forms.CommandButton.1" "btnClear" 108 460 90 28 @{
        Caption = "Clear"
    } | Out-Null

    Add-Control $d "Forms.CommandButton.1" "btnAttach" 204 460 90 28 @{
        Caption = "Attach"
    } | Out-Null

    Add-Control $d "Forms.CommandButton.1" "btnSettings" 300 460 90 28 @{
        Caption = "Settings"
    } | Out-Null

    Add-Control $d "Forms.CommandButton.1" "btnClose" 422 460 90 28 @{
        Caption = "Close"
    } | Out-Null

    Add-Control $d "Forms.Label.1" "lblAttachment" 12 418 500 14 @{
        Caption = ""
    } | Out-Null

    Add-FormCode -Component $frmChat -FormCodePath (Join-Path $InputDir "frmChat.frm")

    $frmSettings = $vb.VBComponents.Add(3)
    $frmSettings.Name = "frmSettings"
    $d2 = $frmSettings.Designer
    try { $d2.Caption = "AI Assistant Settings" } catch {}
    Set-FormSize -Component $frmSettings -Designer $d2 -Width 500 -Height 340

    Add-Control $d2 "Forms.Label.1" "lblDeepSeek" 12 16 120 18 @{
        Caption = "DeepSeek API key"
    } | Out-Null

    Add-Control $d2 "Forms.TextBox.1" "txtDeepSeekKey" 136 14 313 22 @{} | Out-Null

    Add-Control $d2 "Forms.Label.1" "lblOpenRouter" 12 48 120 18 @{
        Caption = "OpenRouter API key"
    } | Out-Null

    Add-Control $d2 "Forms.TextBox.1" "txtOpenRouterKey" 136 46 313 22 @{} | Out-Null

    Add-Control $d2 "Forms.Label.1" "lblOpenAI" 12 80 120 18 @{
        Caption = "OpenAI API key"
    } | Out-Null

    Add-Control $d2 "Forms.TextBox.1" "txtOpenAIKey" 136 78 313 22 @{} | Out-Null

    Add-Control $d2 "Forms.Label.1" "lblLMStudioIP" 12 112 90 18 @{
        Caption = "LM Studio IP"
    } | Out-Null

    Add-Control $d2 "Forms.TextBox.1" "txtLMStudioIP" 106 110 126 22 @{} | Out-Null

    Add-Control $d2 "Forms.Label.1" "lblLMStudioPort" 240 112 36 18 @{
        Caption = "Port"
    } | Out-Null

    Add-Control $d2 "Forms.TextBox.1" "txtLMStudioPort" 280 110 169 22 @{} | Out-Null

    Add-Control $d2 "Forms.Label.1" "lblLMStudioModel" 12 146 120 18 @{
        Caption = "LM Studio model"
    } | Out-Null

    Add-Control $d2 "Forms.ComboBox.1" "cmbLMStudioModel" 136 144 313 22 @{} | Out-Null

    Add-Control $d2 "Forms.Label.1" "lblResponseLanguage" 12 178 120 18 @{
        Caption = "Response language"
    } | Out-Null

    Add-Control $d2 "Forms.ComboBox.1" "cmbResponseLanguage" 136 176 313 22 @{
        Style = 2
    } | Out-Null

    Add-Control $d2 "Forms.Label.1" "lblLMStatus" 12 210 437 18 @{
        Caption = ""
    } | Out-Null

    Add-Control $d2 "Forms.CommandButton.1" "btnRefreshModels" 12 242 130 28 @{
        Caption = "Refresh Models"
    } | Out-Null

    Add-Control $d2 "Forms.CommandButton.1" "btnSave" 263 242 84 28 @{
        Caption = "Save"
    } | Out-Null

    Add-Control $d2 "Forms.CommandButton.1" "btnCancel" 355 242 84 28 @{
        Caption = "Cancel"
    } | Out-Null

    Add-FormCode -Component $frmSettings -FormCodePath (Join-Path $InputDir "frmSettings.frm")

    $twbCode = Get-VbaBody -Path (Join-Path $InputDir "ThisWorkbook.cls")
    if (-not [string]::IsNullOrWhiteSpace($twbCode)) {
        $vb.VBComponents.Item("ThisWorkbook").CodeModule.AddFromString($twbCode)
    }

    $sheetCode = Get-VbaBody -Path (Join-Path $InputDir "Sheet1.cls")
    if (-not [string]::IsNullOrWhiteSpace($sheetCode)) {
        $sheetCodeName = $wb.Worksheets.Item(1).CodeName
        $vb.VBComponents.Item($sheetCodeName).CodeModule.AddFromString($sheetCode)
    }

    $xlOpenXMLWorkbookMacroEnabled = 52
    $wb.SaveAs($outputFull, $xlOpenXMLWorkbookMacroEnabled)
    Write-Output "Created: $outputFull"

    if ($UpdateLatestAlias) {
        $latestName = "AI_Assistant_latest_{0}.xlsm" -f $Track.ToLowerInvariant()
        $latestPath = Join-Path (Split-Path -Parent $outputFull) $latestName
    }
}
finally {
    if ($wb) {
        $wb.Close($false) | Out-Null
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    if ($excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($UpdateLatestAlias -and -not [string]::IsNullOrWhiteSpace($latestPath)) {
    Copy-Item -LiteralPath $outputFull -Destination $latestPath -Force
    Write-Output "Updated: $latestPath"
}
