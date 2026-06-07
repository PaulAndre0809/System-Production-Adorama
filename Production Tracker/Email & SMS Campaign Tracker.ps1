[CmdletBinding()]
param(
    [switch]$Visible,
    [switch]$SkipQa
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$baseName = "Email & SMS Campaign Tracker"
$workbookPath = Join-Path $PSScriptRoot "$baseName.xlsm"
$modulePath = Join-Path $PSScriptRoot "$baseName.bas"
$qaPath = Join-Path $PSScriptRoot "$baseName.py"
$pythonPath = Join-Path (Split-Path $PSScriptRoot -Parent) ".venv\Scripts\python.exe"
$securityPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$accessProperty = "AccessVBOM"
$transactionId = [guid]::NewGuid().ToString("N")
$workingPath = Join-Path $PSScriptRoot ".$baseName.$transactionId.working.xlsm"
$rollbackPath = Join-Path $PSScriptRoot ".$baseName.$transactionId.rollback.xlsm"
$failedPath = Join-Path $PSScriptRoot ".$baseName.$transactionId.failed.xlsm"

$workbookOpenCode = @'
Private Sub Workbook_Open()
    Dim priorEvents As Boolean

    priorEvents = Application.EnableEvents
    On Error GoTo SafeExit
    Application.EnableEvents = False
    ThisWorkbook.ForceFullCalculation = False
    modEmailProductionTracker.RefreshProductionStatus
    modEmailProductionTracker.UpdateCalendarTabs
SafeExit:
    Application.EnableEvents = priorEvents
End Sub
'@

$campaignSheetCode = @'
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim lo As ListObject
    Dim priorEvents As Boolean

    On Error GoTo SafeExit
    If Me.ListObjects.Count = 0 Then Exit Sub
    Set lo = Me.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub

    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    Cancel = modEmailProductionTracker.ToggleInventoryChecklist(Target)
SafeExit:
    Application.EnableEvents = priorEvents
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim lo As ListObject
    Dim priorEvents As Boolean

    On Error GoTo SafeExit
    If Me.ListObjects.Count = 0 Then Exit Sub
    Set lo = Me.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub

    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    modEmailProductionTracker.HandleCampaignChange Me, Target
SafeExit:
    Application.EnableEvents = priorEvents
End Sub
'@

function Release-ComObject {
    param($Object)

    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Set-CodeModule {
    param(
        [Parameter(Mandatory)]$CodeModule,
        [Parameter(Mandatory)][string]$Code
    )

    if ($CodeModule.CountOfLines -gt 0) {
        $CodeModule.DeleteLines(1, $CodeModule.CountOfLines)
    }
    $CodeModule.AddFromString($Code)
}

function Get-Worksheet {
    param(
        [Parameter(Mandatory)]$Workbook,
        [Parameter(Mandatory)][string]$Name
    )

    try {
        return $Workbook.Worksheets.Item($Name)
    } catch {
        return $null
    }
}

function Test-WorksheetHasDataRows {
    param([Parameter(Mandatory)]$Worksheet)

    $used = $Worksheet.UsedRange
    try {
        $lastRow = $used.Row + $used.Rows.Count - 1
        $lastColumn = $used.Column + $used.Columns.Count - 1
        if ($lastRow -lt 2) {
            return $false
        }

        $dataRange = $Worksheet.Range(
            $Worksheet.Cells(2, 1),
            $Worksheet.Cells($lastRow, $lastColumn)
        )
        try {
            return $Worksheet.Application.WorksheetFunction.CountA($dataRange) -gt 0
        } finally {
            Release-ComObject $dataRange
        }
    } finally {
        Release-ComObject $used
    }
}

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw "Close all Excel windows before running the deployment."
}
if (-not (Test-Path -LiteralPath $workbookPath)) {
    throw "Workbook not found: $workbookPath"
}
if (-not (Test-Path -LiteralPath $modulePath)) {
    throw "VBA source not found: $modulePath"
}
if (-not $SkipQa) {
    if (-not (Test-Path -LiteralPath $qaPath)) {
        throw "QA utility not found: $qaPath"
    }
    if (-not (Test-Path -LiteralPath $pythonPath)) {
        throw "Python environment not found: $pythonPath"
    }
}

Copy-Item -LiteralPath $workbookPath -Destination $workingPath

$existingSetting = Get-ItemProperty `
    -LiteralPath $securityPath `
    -Name $accessProperty `
    -ErrorAction SilentlyContinue
$hadSetting = $null -ne $existingSetting
$oldAccessValue = if ($hadSetting) { $existingSetting.$accessProperty } else { $null }

$excel = $null
$workbook = $null
$dummyWorkbook = $null
$components = $null
$replacementCommitted = $false
$deploymentSucceeded = $false
$oldCalculation = $null
$oldCalculateBeforeSave = $null

try {
    New-Item -ItemType Directory -Path $securityPath -Force | Out-Null
    Set-ItemProperty `
        -LiteralPath $securityPath `
        -Name $accessProperty `
        -Type DWord `
        -Value 1

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $oldCalculateBeforeSave = $excel.CalculateBeforeSave

    $dummyWorkbook = $excel.Workbooks.Add()
    $oldCalculation = $excel.Calculation
    $excel.Calculation = -4135

    Write-Host "Opening transactional workbook..."
    $workbook = $excel.Workbooks.Open($workingPath, 0, $false)
    $excel.CalculateBeforeSave = $false

    $emailSheet = Get-Worksheet $workbook "Email Campaigns"
    $productionSheet = Get-Worksheet $workbook "Production Inventory"
    if ($null -eq $emailSheet -and $null -ne $productionSheet) {
        $productionSheet.Name = "Email Campaigns"
        $emailSheet = $productionSheet
        $productionSheet = $null
    } elseif ($null -eq $emailSheet) {
        throw "Neither Production Inventory nor Email Campaigns exists."
    } elseif ($null -ne $productionSheet) {
        throw "Both Production Inventory and Email Campaigns exist; deployment stopped to protect data."
    }

    $smsSheet = Get-Worksheet $workbook "SMS Campaigns"
    if ($null -eq $smsSheet) {
        $smsSheet = $workbook.Worksheets.Add()
        $smsSheet.Name = "SMS Campaigns"
    }

    foreach ($sheetName in @("Combined Campaigns")) {
        $legacySheet = Get-Worksheet $workbook $sheetName
        if ($null -ne $legacySheet) {
            if (Test-WorksheetHasDataRows $legacySheet) {
                Write-Warning "Preserving '$sheetName' because it contains data."
            } else {
                $legacySheet.Visible = -1
                $legacySheet.Delete()
            }
            Release-ComObject $legacySheet
        }
    }

    $components = $workbook.VBProject.VBComponents
    for ($index = $components.Count; $index -ge 1; $index--) {
        $component = $components.Item($index)
        try {
            if ($component.Type -eq 1) {
                $components.Remove($component)
            }
        } finally {
            Release-ComObject $component
        }
    }

    Write-Host "Saving VBA purge phase..."
    $workbook.Save()
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $null
    Release-ComObject $components
    $components = $null

    Write-Host "Reopening for VBA injection..."
    $workbook = $excel.Workbooks.Open($workingPath, 0, $false)
    $components = $workbook.VBProject.VBComponents

    $moduleCode = Get-Content -LiteralPath $modulePath -Raw
    $moduleCode = $moduleCode -replace '^\s*Attribute VB_Name = "[^"]+"\r?\n', ""

    $standardModule = $components.Add(1)
    try {
        $standardModule.Name = "modEmailProductionTracker"
        Set-CodeModule -CodeModule $standardModule.CodeModule -Code $moduleCode
    } finally {
        Release-ComObject $standardModule
    }

    $thisWorkbookComponent = $components.Item("ThisWorkbook")
    try {
        Set-CodeModule `
            -CodeModule $thisWorkbookComponent.CodeModule `
            -Code $workbookOpenCode
    } finally {
        Release-ComObject $thisWorkbookComponent
    }

    Release-ComObject $emailSheet
    $emailSheet = $workbook.Worksheets.Item("Email Campaigns")
    $sheetComponent = $components.Item($emailSheet.CodeName)
    try {
        Set-CodeModule `
            -CodeModule $sheetComponent.CodeModule `
            -Code $campaignSheetCode
    } finally {
        Release-ComObject $sheetComponent
    }

    Release-ComObject $smsSheet
    $smsSheet = $workbook.Worksheets.Item("SMS Campaigns")
    $sheetComponent = $components.Item($smsSheet.CodeName)
    try {
        Set-CodeModule `
            -CodeModule $sheetComponent.CodeModule `
            -Code $campaignSheetCode
    } finally {
        Release-ComObject $sheetComponent
    }

    $compileControl = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compileControl) {
        throw "Excel's VBA compile command is unavailable."
    }
    if ($compileControl.Enabled) {
        $compileControl.Execute()
    }
    Release-ComObject $compileControl

    $macroPrefix = "'$($workbook.Name)'!modEmailProductionTracker."
    Write-Host "Applying workbook configuration..."
    $excel.Run($macroPrefix + "ApplyAllConfigurations")
    $validationResult = [string]$excel.Run(
        $macroPrefix + "ValidateWorkbookConfiguration"
    )
    if ($validationResult -ne "OK") {
        throw $validationResult
    }

    $workbook.ForceFullCalculation = $false
    Write-Host "Calculating configured worksheets..."
    foreach ($sheet in $workbook.Worksheets) {
        $sheet.Calculate()
        Release-ComObject $sheet
    }
    Write-Host "Switching workbook to automatic calculation..."
    $excel.Calculation = -4105
    Write-Host "Saving final workbook..."
    $workbook.Save()
    $excel.CalculateBeforeSave = $oldCalculateBeforeSave
    $workbook.Close($false)
    Release-ComObject $workbook
    $workbook = $null

    $dummyWorkbook.Close($false)
    Release-ComObject $dummyWorkbook
    $dummyWorkbook = $null
    $excel.Quit()
    Release-ComObject $excel
    $excel = $null

    [IO.File]::Replace($workingPath, $workbookPath, $rollbackPath, $true)
    $replacementCommitted = $true

    if (-not $SkipQa) {
        & $pythonPath $qaPath --qa --workbook $workbookPath
        if ($LASTEXITCODE -ne 0) {
            throw "Post-deployment QA failed with exit code $LASTEXITCODE."
        }
    }

    $deploymentSucceeded = $true
    Write-Host "Deployment and validation completed successfully." -ForegroundColor Green
    Write-Host "Workbook: $workbookPath"
    Write-Host "Validation: $validationResult"
} catch {
    if ($replacementCommitted -and (Test-Path -LiteralPath $rollbackPath)) {
        [IO.File]::Replace($rollbackPath, $workbookPath, $failedPath, $true)
        Remove-Item -LiteralPath $failedPath -Force -ErrorAction SilentlyContinue
        $replacementCommitted = $false
    }
    throw
} finally {
    if ($null -ne $dummyWorkbook) {
        try { $dummyWorkbook.Close($false) } catch {}
        Release-ComObject $dummyWorkbook
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }
    Release-ComObject $components

    if ($null -ne $excel) {
        try {
            if ($null -ne $oldCalculation) {
                $excel.Calculation = $oldCalculation
            }
            if ($null -ne $oldCalculateBeforeSave) {
                $excel.CalculateBeforeSave = $oldCalculateBeforeSave
            }
            $excel.Quit()
        } catch {}
        Release-ComObject $excel
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($hadSetting) {
        Set-ItemProperty `
            -LiteralPath $securityPath `
            -Name $accessProperty `
            -Type DWord `
            -Value $oldAccessValue
    } else {
        Remove-ItemProperty `
            -LiteralPath $securityPath `
            -Name $accessProperty `
            -ErrorAction SilentlyContinue
    }

    if (Test-Path -LiteralPath $workingPath) {
        Remove-Item -LiteralPath $workingPath -Force
    }
    if ($deploymentSucceeded -and (Test-Path -LiteralPath $rollbackPath)) {
        Remove-Item -LiteralPath $rollbackPath -Force
    }
}
