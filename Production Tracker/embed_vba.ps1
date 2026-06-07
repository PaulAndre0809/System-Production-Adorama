[CmdletBinding()]
param(
    [switch]$Visible
)

$ErrorActionPreference = "Stop"

$preferredWorkbook = "Email & SMS Campaign Tracker.xlsm"
$legacyWorkbook = "Email_Production_Inventory_Tracker_UI.xlsm"
$preferredModule = "Email & SMS Campaign Tracker.bas"
$legacyModule = "Email_Production_Inventory_Tracker_UI.bas"

if (Test-Path (Join-Path $PSScriptRoot $preferredWorkbook)) {
    $workbookPath = Join-Path $PSScriptRoot $preferredWorkbook
} else {
    $workbookPath = Join-Path $PSScriptRoot $legacyWorkbook
}

if (Test-Path (Join-Path $PSScriptRoot $preferredModule)) {
    $modulePath = Join-Path $PSScriptRoot $preferredModule
} else {
    $modulePath = Join-Path $PSScriptRoot $legacyModule
}
$securityPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Security"
$accessProperty = "AccessVBOM"
$workingPath = Join-Path $env:TEMP (
    "ProductionTracker_Embed_" + [guid]::NewGuid().ToString("N") + ".xlsm"
)

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw "Close all Excel windows before embedding VBA."
}

if (-not (Test-Path -LiteralPath $workbookPath)) {
    throw "Workbook not found: $workbookPath"
}

if (-not (Test-Path -LiteralPath $modulePath)) {
    throw "VBA module not found: $modulePath"
}

Copy-Item -LiteralPath $workbookPath -Destination $workingPath

$existingSetting = Get-ItemProperty `
    -LiteralPath $securityPath `
    -Name $accessProperty `
    -ErrorAction SilentlyContinue
$hadSetting = $null -ne $existingSetting
$oldAccessValue = if ($hadSetting) {
    $existingSetting.$accessProperty
} else {
    $null
}

$excel = $null
$workbook = $null

$workbookOpenCode = @'
Private Sub Workbook_Open()
    On Error Resume Next
    modEmailProductionTracker.RefreshProductionStatus
    modEmailProductionTracker.UpdateCalendarTabs
End Sub
'@

$inventorySheetCode = @'
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo SafeExit
    If Intersect(Target, Me.ListObjects("ProductionInventoryTable").DataBodyRange) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    Call modEmailProductionTracker.ToggleInventoryChecklist(Target)
SafeExit:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    Dim lo As ListObject
    Set lo = Me.ListObjects("ProductionInventoryTable")
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    modEmailProductionTracker.HandleInventoryChange Target
SafeExit:
    Application.EnableEvents = True
End Sub
'@

function Set-CodeModule {
    param(
        [Parameter(Mandatory)]
        $CodeModule,

        [Parameter(Mandatory)]
        [string]$Code
    )

    if ($CodeModule.CountOfLines -gt 0) {
        $CodeModule.DeleteLines(1, $CodeModule.CountOfLines)
    }
    $CodeModule.AddFromString($Code)
}

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

    $workbook = $excel.Workbooks.Open($workingPath)
    $components = $workbook.VBProject.VBComponents

    for ($index = $components.Count; $index -ge 1; $index--) {
        $component = $components.Item($index)
        if ($component.Type -eq 1) {
            $components.Remove($component)
        }
    }

    $workbook.Save()
    $workbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    $workbook = $null

    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($components)
    $components = $null

    $workbook = $excel.Workbooks.Open($workingPath)
    $components = $workbook.VBProject.VBComponents

    $moduleCode = Get-Content -LiteralPath $modulePath -Raw
    $moduleCode = $moduleCode -replace `
        '^\s*Attribute VB_Name = "[^"]+"\r?\n', ""

    $importedModule = $components.Add(1)
    $importedModule.Name = "modEmailProductionTracker"
    Set-CodeModule `
        -CodeModule $importedModule.CodeModule `
        -Code $moduleCode

    if ($importedModule.Name -ne "modEmailProductionTracker") {
        throw "Unable to normalize imported VBA module name."
    }

    Set-CodeModule `
        -CodeModule $components.Item("ThisWorkbook").CodeModule `
        -Code $workbookOpenCode

    $inventorySheet = $workbook.Worksheets.Item("Production Inventory")
    Set-CodeModule `
        -CodeModule $components.Item($inventorySheet.CodeName).CodeModule `
        -Code $inventorySheetCode

    $compileControl = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -ne $compileControl -and $compileControl.Enabled) {
        $compileControl.Execute()
    }

    $excel.EnableEvents = $false
    $macroName = "'$($workbook.Name)'!" +
        "modEmailProductionTracker.ApplyAllConfigurations"
    $excel.Run($macroName)

    $validationName = "'$($workbook.Name)'!" +
        "modEmailProductionTracker.ValidateWorkbookConfiguration"
    $validationResult = [string]$excel.Run($validationName)
    if ($validationResult -ne "OK") {
        throw $validationResult
    }

    $workbook.Save()

    $workbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    $workbook = $null

    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    $excel = $null

    Copy-Item `
        -LiteralPath $workingPath `
        -Destination $workbookPath `
        -Force

    Write-Host "VBA embedded and workbook QA passed." -ForegroundColor Green
    Write-Host "Validation: $validationResult"
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
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
}
