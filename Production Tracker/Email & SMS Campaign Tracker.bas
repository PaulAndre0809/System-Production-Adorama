Attribute VB_Name = "modEmailProductionTracker"
Option Explicit

Public Const SH_INVENTORY As String = "Production Inventory"
Public Const SH_DASHBOARD As String = "Dashboard"
Public Const SH_LOG As String = "Automation Log"

Private Const TBL_INVENTORY As String = "ProductionInventoryTable"
Private Const TBL_DASHBOARD As String = "DashboardWorkTable"
Private Const FIRST_DASHBOARD_ROW As Long = 11
Private Const CALENDAR_INVENTORY_SHEET As String = "Production Inventory"
Private Const CALENDAR_DASHBOARD_SHEET As String = "Dashboard"
Private Const CALENDAR_TABLE As String = "ProductionInventoryTable"
Private Const FIRST_CALENDAR_ROW As Long = 6
Private Const LAST_CALENDAR_ROW As Long = 11

'Run this procedure once after replacing the old modEmailProductionTracker module.
Public Sub MigrateProductionInventoryStructure()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim deleteHeaders As Variant
    Dim checklistHeaders As Variant
    Dim item As Variant
    Dim backupPath As String
    Dim brokenBefore As Long
    Dim brokenAfter As Long
    Dim oldCalculation As XlCalculation
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldStatusBar As Variant
    Dim completed As Boolean
    Dim failureText As String

    Set wb = ThisWorkbook

    oldCalculation = Application.Calculation
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar

    On Error GoTo MigrationFailed

    Set ws = wb.Worksheets(SH_INVENTORY)
    Set lo = ws.ListObjects(TBL_INVENTORY)

    If wb.ProtectStructure Then
        Err.Raise vbObjectError + 1000, , _
            "The workbook structure is protected. Unprotect it before running the migration."
    End If

    If wb.ReadOnly Then
        Err.Raise vbObjectError + 1003, , _
            "The workbook is read-only. Reopen an editable copy before running the migration."
    End If

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 1001, , _
            "The Production Inventory sheet is protected. Unprotect it before running the migration."
    End If

    RequireTableColumn lo, "Send Date"
    RequireTableColumn lo, "Campaign Name"
    RequireTableColumn lo, "Current Stage"
    RequireTableColumn lo, "Owner"
    RequireTableColumn lo, "Delivered"
    RequireTableColumn lo, "Last Updated"

    backupPath = CreateBackupCopy(wb)
    brokenBefore = CountBrokenReferences(wb)

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Updating Production Inventory structure..."

    deleteHeaders = Array( _
        "Campaign ID", _
        "Missing / Blocker", _
        "Next Action", _
        "SKU Status", _
        "Brief Status", _
        "Design Status", _
        "Builder Status", _
        "Segment Status", _
        "QA Status", _
        "Schedule Status", _
        "Risk Level")

    For Each item In deleteHeaders
        DeleteTableColumnIfPresent lo, CStr(item)
    Next item

    checklistHeaders = Array( _
        "Campaign Name and UTM Parameter (Source Code)", _
        "Creative Brief, SL & PH", _
        "SKUs", _
        "In-Design", _
        "Build, QA", _
        "Route", _
        "Approval", _
        "Segments")

    AddChecklistColumns lo, checklistHeaders, "Owner"

    For Each item In checklistHeaders
        ConfigureChecklistColumn lo, CStr(item)
    Next item

    ' Ensure Owner is plain text and ensure Last Updated By column exists
    ConfigureOwnerColumn lo
    EnsureUpdatedByColumn lo

    FormatSendDateColumn lo
    ApplyCalculatedColumns lo
    RefreshDashboard

    Application.CalculateFull
    brokenAfter = CountBrokenReferences(wb)

    If brokenAfter > brokenBefore Then
        Err.Raise vbObjectError + 1002, , _
            "The migration introduced " & (brokenAfter - brokenBefore) & _
            " new #REF! reference(s). Close without saving and restore the backup."
    End If

    LogAction "MigrateProductionInventoryStructure", _
        "Production Inventory migrated. Backup: " & backupPath

    completed = True

MigrationExit:
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts
    Application.StatusBar = oldStatusBar

    If completed Then
        MsgBox _
            "Production Inventory migration completed." & vbCrLf & vbCrLf & _
            "Backup created:" & vbCrLf & backupPath & vbCrLf & vbCrLf & _
            "Review the workbook, then save it as XLSM.", _
            vbInformation, "Migration Complete"
    Else
        If Len(backupPath) > 0 Then
            MsgBox _
                "Migration stopped: " & failureText & vbCrLf & vbCrLf & _
                "Do not save the current workbook. Backup:" & vbCrLf & backupPath, _
                vbCritical, "Migration Stopped"
        Else
            MsgBox _
                "Migration stopped before a backup was created: " & _
                failureText, vbCritical, "Migration Stopped"
        End If
    End If
    Exit Sub

MigrationFailed:
    failureText = Err.Description
    Resume MigrationExit
End Sub

'Preserves the existing public entry point while removing fixed column letters.
Public Sub RefreshProductionStatus()
    Dim lo As ListObject
    Dim oldCalculation As XlCalculation
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    oldCalculation = Application.Calculation
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    On Error GoTo RefreshFailed

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set lo = GetInventoryTable()
    ApplyCalculatedColumns lo
    RefreshDashboard
    LogAction "RefreshProductionStatus", _
        "Production Inventory and Dashboard refreshed"

RefreshExit:
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If failureNumber <> 0 Then
        Err.Raise failureNumber, "RefreshProductionStatus", failureDescription
    End If
    Exit Sub

RefreshFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume RefreshExit
End Sub

'Public wrapper for automation scripts to apply configurations without passing ListObjects
Public Sub ApplyAllConfigurations()
    Dim lo As ListObject
    Dim checklistHeaders As Variant
    Dim item As Variant
    Dim validationResult As String
    Dim oldCalculation As XlCalculation
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    oldCalculation = Application.Calculation
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    On Error GoTo ApplyFailed
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set lo = GetInventoryTable()

    checklistHeaders = Array( _
        "Campaign Name and UTM Parameter (Source Code)", _
        "Creative Brief, SL & PH", _
        "SKUs", _
        "In-Design", _
        "Build, QA", _
        "Route", _
        "Approval", _
        "Segments")

    For Each item In checklistHeaders
        ConfigureChecklistColumn lo, CStr(item)
    Next item

    ConfigureOwnerColumn lo
    EnsureUpdatedByColumn lo
    FormatSendDateColumn lo
    ApplyCalculatedColumns lo
    RefreshDashboard
    UpdateCalendarTabs

    validationResult = ValidateWorkbookConfiguration()
    If validationResult <> "OK" Then
        Err.Raise vbObjectError + 1090, , validationResult
    End If

    LogAction "ApplyAllConfigurations", "All configurations applied successfully"

ApplyExit:
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If failureNumber <> 0 Then
        Err.Raise failureNumber, "ApplyAllConfigurations", failureDescription
    End If
    Exit Sub

ApplyFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume ApplyExit
End Sub

'Retained for compatibility with any existing callers or UserForms.
Public Function GenerateCampaignID( _
    ByVal sendDate As Variant, _
    ByVal campaignName As String) As String

    Dim cleanName As String

    cleanName = UCase$(campaignName)
    cleanName = Replace(cleanName, "Sample | ", "")
    cleanName = Replace(cleanName, " ", "-")
    cleanName = Replace(cleanName, "/", "-")
    cleanName = Replace(cleanName, "&", "AND")
    cleanName = Replace(cleanName, "--", "-")

    If IsDate(sendDate) Then
        GenerateCampaignID = _
            "EMAIL-" & Format$(CDate(sendDate), "yyyy-mm-dd") & _
            "-" & Left$(cleanName, 25)
    Else
        GenerateCampaignID = "EMAIL-TBD-" & Left$(cleanName, 25)
    End If
End Function

'Retained with the original signature for compatibility.
Public Function CalculateCurrentStage( _
    ByVal ws As Worksheet, _
    ByVal r As Long) As String

    Dim delivered As Variant
    If Len(Trim$(CStr(ValueByHeader(ws, r, "Campaign Name")))) = 0 Then
        CalculateCurrentStage = vbNullString
        Exit Function
    End If

    delivered = ValueByHeader(ws, r, "Delivered")

    If IsNumeric(delivered) And delivered > 0 Then
        CalculateCurrentStage = "Sent"
    ElseIf Not IsChecked(ValueByHeader( _
        ws, r, "Campaign Name and UTM Parameter (Source Code)")) Then
        CalculateCurrentStage = "Source Code"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "Creative Brief, SL & PH")) Then
        CalculateCurrentStage = "Creative Brief"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "SKUs")) Then
        CalculateCurrentStage = "Waiting for SKUs"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "In-Design")) Then
        CalculateCurrentStage = "With Design"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "Build, QA")) Then
        CalculateCurrentStage = "Build / QA"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "Route")) Then
        CalculateCurrentStage = "Routing"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "Approval")) Then
        CalculateCurrentStage = "Awaiting Approval"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "Segments")) Then
        CalculateCurrentStage = "Segments"
    ElseIf Not HasText(ValueByHeader(ws, r, "Jira Link")) Or _
        Not HasText(ValueByHeader(ws, r, "ClickUp Link")) Or _
        Not HasText(ValueByHeader(ws, r, "Bluecore Link")) Then
        CalculateCurrentStage = "Links Pending"
    ElseIf Not HasText(ValueByHeader(ws, r, "Est. Audience")) Then
        CalculateCurrentStage = "Ready to Schedule"
    Else
        CalculateCurrentStage = "Scheduled"
    End If
End Function

'Retained with the original signature for compatibility. Risk is now derived,
'because the stored Risk Level column is intentionally removed.
Public Function CalculateRiskLevel( _
    ByVal ws As Worksheet, _
    ByVal r As Long) As String

    Dim sendDate As Variant
    Dim currentStage As String
    Dim daysLeft As Long

    sendDate = ValueByHeader(ws, r, "Send Date")
    currentStage = CalculateCurrentStage(ws, r)

    If currentStage = "Sent" Or currentStage = "Cancelled" Then
        CalculateRiskLevel = "Complete"
        Exit Function
    End If

    If currentStage = "Scheduled" Or currentStage = "Ready to Schedule" Then
        CalculateRiskLevel = "Ready"
        Exit Function
    End If

    If Not IsDate(sendDate) Then
        CalculateRiskLevel = "At Risk"
        Exit Function
    End If

    daysLeft = DateDiff("d", Date, CDate(sendDate))

    If daysLeft <= 0 Then
        CalculateRiskLevel = "Critical"
    ElseIf daysLeft = 1 Then
        CalculateRiskLevel = "High Risk"
    ElseIf daysLeft <= 2 Then
        CalculateRiskLevel = "At Risk"
    Else
        CalculateRiskLevel = "On Track"
    End If
End Function

Public Sub RefreshDashboard()
    Dim wsD As Worksheet
    Dim wsI As Worksheet
    Dim lo As ListObject
    Dim dashboardTable As ListObject
    Dim i As Long
    Dim r As Long
    Dim matchCount As Long
    Dim displayRows As Long
    Dim displayIndex As Long
    Dim oldLastDashboardRow As Long
    Dim lastDashboardRow As Long
    Dim summaryRow As Long
    Dim dashboardValues() As Variant
    Dim jiraLinks() As String
    Dim clickUpLinks() As String
    Dim bluecoreLinks() As String
    Dim weekStart As Date
    Dim weekEnd As Date
    Dim campaignName As String
    Dim sendDate As Variant
    Dim checkedMark As String
    Dim campaignRef As String
    Dim stageRef As String
    Dim sendDateRef As String
    Dim sourceCodeRef As String
    Dim briefRef As String
    Dim buildQaRef As String
    Dim approvalRef As String

    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set wsI = ThisWorkbook.Worksheets(SH_INVENTORY)
    Set lo = GetInventoryTable()
    Set dashboardTable = wsD.ListObjects(TBL_DASHBOARD)
    checkedMark = CheckedSymbol()

    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1070, , _
            "Production Inventory has no data rows."
    End If

    weekStart = Date - Weekday(Date, vbMonday) + 1
    weekEnd = weekStart + 13

    For i = 1 To lo.ListRows.Count
        r = lo.DataBodyRange.Rows(i).Row
        campaignName = Trim$(TextValue( _
            ValueByHeader(wsI, r, "Campaign Name")))
        sendDate = ValueByHeader(wsI, r, "Send Date")

        If Len(campaignName) > 0 And Not IsError(sendDate) Then
            If IsDate(sendDate) Then
                If DateValue(CDate(sendDate)) >= weekStart And _
                    DateValue(CDate(sendDate)) <= weekEnd Then
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next i

    displayRows = matchCount
    If displayRows = 0 Then displayRows = 1

    oldLastDashboardRow = dashboardTable.Range.Row + _
        dashboardTable.Range.Rows.Count - 1
    If Not dashboardTable.DataBodyRange Is Nothing Then
        On Error Resume Next
        dashboardTable.DataBodyRange.Hyperlinks.Delete
        On Error GoTo 0
        dashboardTable.DataBodyRange.ClearContents
    End If

    wsD.Range( _
        "A" & (oldLastDashboardRow + 5) & _
        ":D" & (oldLastDashboardRow + 9)).ClearContents
    wsD.Range("A115:D119").ClearContents

    lastDashboardRow = FIRST_DASHBOARD_ROW + displayRows - 1
    summaryRow = lastDashboardRow + 5
    dashboardTable.Resize wsD.Range("A10:L" & lastDashboardRow)

    ReDim dashboardValues(1 To displayRows, 1 To 12)
    ReDim jiraLinks(1 To displayRows)
    ReDim clickUpLinks(1 To displayRows)
    ReDim bluecoreLinks(1 To displayRows)

    campaignRef = FullColumnReference(lo, "Campaign Name")
    stageRef = FullColumnReference(lo, "Current Stage")
    sendDateRef = FullColumnReference(lo, "Send Date")
    sourceCodeRef = FullColumnReference( _
        lo, "Campaign Name and UTM Parameter (Source Code)")
    briefRef = FullColumnReference(lo, "Creative Brief, SL & PH")
    buildQaRef = FullColumnReference(lo, "Build, QA")
    approvalRef = FullColumnReference(lo, "Approval")

    wsD.Range("A2").Value = _
        "Daily production view using the new checklist workflow."

    wsD.Range("A4").Value = "Active Work"
    wsD.Range("C4").Value = "Sending Today"
    wsD.Range("E4").Value = "Source Pending"
    wsD.Range("G4").Value = "Brief Pending"
    wsD.Range("I4").Value = "Approval Pending"
    wsD.Range("K4").Value = "Sent"

    wsD.Range("A6").Value = "Not sent or cancelled"
    wsD.Range("C6").Value = "Scheduled for today"
    wsD.Range("E6").Value = "Source code incomplete"
    wsD.Range("G6").Value = "Creative brief incomplete"
    wsD.Range("I6").Value = "Approval incomplete"
    wsD.Range("K6").Value = "Delivered or sent"

    wsD.Range("B3").Formula = "=NOW()"
    wsD.Range("B3").NumberFormat = "MM/DD/YYYY HH:MM"
    If wsD.Columns("B").ColumnWidth < 20 Then
        wsD.Columns("B").ColumnWidth = 20
    End If
    wsD.Range("A5").Formula = _
        "=COUNTIFS(" & campaignRef & ",""""," & _
        stageRef & ",""<>Sent""," & stageRef & ",""<>Cancelled"")"
    wsD.Range("C5").Formula = _
        "=COUNTIFS(" & sendDateRef & ",TODAY()," & campaignRef & ","""")"
    wsD.Range("E5").Formula = PendingCountFormula( _
        campaignRef, sourceCodeRef, checkedMark)
    wsD.Range("G5").Formula = PendingCountFormula( _
        campaignRef, briefRef, checkedMark)
    wsD.Range("I5").Formula = PendingCountFormula( _
        campaignRef, approvalRef, checkedMark)
    wsD.Range("K5").Formula = _
        "=COUNTIFS(" & stageRef & ",""Sent"")"

    wsD.Range("A9").Value = _
        "Rows mirror Production Inventory; checklist cells show completion."

    wsD.Range("A10").Value = "Send Date"
    wsD.Range("B10").Value = "Time"
    wsD.Range("C10").Value = "Campaign"
    wsD.Range("D10").Value = "Type"
    wsD.Range("E10").Value = "Stage"
    wsD.Range("F10").Value = "Owner"
    wsD.Range("G10").Value = "SKUs"
    wsD.Range("H10").Value = "Build, QA"
    wsD.Range("I10").Value = "Approval"
    wsD.Range("J10").Value = "Jira"
    wsD.Range("K10").Value = "ClickUp"
    wsD.Range("L10").Value = "Bluecore"

    For i = 1 To lo.ListRows.Count
        r = lo.DataBodyRange.Rows(i).Row
        campaignName = Trim$(TextValue( _
            ValueByHeader(wsI, r, "Campaign Name")))
        sendDate = ValueByHeader(wsI, r, "Send Date")

        If Len(campaignName) > 0 And Not IsError(sendDate) Then
            If IsDate(sendDate) Then
                If DateValue(CDate(sendDate)) >= weekStart And _
                    DateValue(CDate(sendDate)) <= weekEnd Then
                    displayIndex = displayIndex + 1
                    dashboardValues(displayIndex, 1) = _
                        DateValue(CDate(sendDate))
                    dashboardValues(displayIndex, 2) = _
                        ValueByHeader(wsI, r, "Send Time")
                    dashboardValues(displayIndex, 3) = campaignName
                    dashboardValues(displayIndex, 4) = _
                        ValueByHeader(wsI, r, "Campaign Type")
                    dashboardValues(displayIndex, 5) = _
                        CalculateCurrentStage(wsI, r)
                    dashboardValues(displayIndex, 6) = _
                        ValueByHeader(wsI, r, "Owner")
                    dashboardValues(displayIndex, 7) = _
                        ValueByHeader(wsI, r, "SKUs")
                    dashboardValues(displayIndex, 8) = _
                        ValueByHeader(wsI, r, "Build, QA")
                    dashboardValues(displayIndex, 9) = _
                        ValueByHeader(wsI, r, "Approval")
                    jiraLinks(displayIndex) = TextValue( _
                        ValueByHeader(wsI, r, "Jira Link"))
                    clickUpLinks(displayIndex) = TextValue( _
                        ValueByHeader(wsI, r, "ClickUp Link"))
                    bluecoreLinks(displayIndex) = TextValue( _
                        ValueByHeader(wsI, r, "Bluecore Link"))
                End If
            End If
        End If
    Next i

    dashboardTable.DataBodyRange.Value2 = dashboardValues

    For i = 1 To matchCount
        AddDashboardLink wsD.Cells(FIRST_DASHBOARD_ROW + i - 1, "J"), _
            jiraLinks(i), "Jira"
        AddDashboardLink wsD.Cells(FIRST_DASHBOARD_ROW + i - 1, "K"), _
            clickUpLinks(i), "ClickUp"
        AddDashboardLink wsD.Cells(FIRST_DASHBOARD_ROW + i - 1, "L"), _
            bluecoreLinks(i), "Bluecore"
    Next i

    wsD.Range("A11:A" & lastDashboardRow).NumberFormat = "MM/DD/YYYY"
    wsD.Range("B11:B" & lastDashboardRow).NumberFormat = "h:mm AM/PM"
    wsD.Range("G11:I" & lastDashboardRow).Font.Name = "Segoe UI Symbol"
    wsD.Range("G11:I" & lastDashboardRow).HorizontalAlignment = xlCenter

    wsD.Cells(summaryRow, "A").Value = "Today's Sends"
    wsD.Cells(summaryRow, "B").Value = "Send Date = today"
    wsD.Cells(summaryRow, "C").Value = "Confirm scheduled/send status"
    wsD.Cells(summaryRow, "D").Formula = _
        "=COUNTIFS(" & sendDateRef & ",TODAY()," & campaignRef & ","""")"

    wsD.Cells(summaryRow + 1, "A").Value = "Source Code Pending"
    wsD.Cells(summaryRow + 1, "B").Value = "Source-code checkbox is open"
    wsD.Cells(summaryRow + 1, "C").Value = _
        "Finish campaign name and UTM parameters"
    wsD.Cells(summaryRow + 1, "D").Formula = PendingCountFormula( _
        campaignRef, sourceCodeRef, checkedMark)

    wsD.Cells(summaryRow + 2, "A").Value = "Creative Brief Pending"
    wsD.Cells(summaryRow + 2, "B").Value = "Creative checkbox is open"
    wsD.Cells(summaryRow + 2, "C").Value = _
        "Complete brief, subject line, and preheader"
    wsD.Cells(summaryRow + 2, "D").Formula = PendingCountFormula( _
        campaignRef, briefRef, checkedMark)

    wsD.Cells(summaryRow + 3, "A").Value = "Build / QA Pending"
    wsD.Cells(summaryRow + 3, "B").Value = "Build/QA checkbox is open"
    wsD.Cells(summaryRow + 3, "C").Value = "Complete build and QA"
    wsD.Cells(summaryRow + 3, "D").Formula = PendingCountFormula( _
        campaignRef, buildQaRef, checkedMark)

    wsD.Cells(summaryRow + 4, "A").Value = "Approval Pending"
    wsD.Cells(summaryRow + 4, "B").Value = "Approval checkbox is open"
    wsD.Cells(summaryRow + 4, "C").Value = "Obtain final approval"
    wsD.Cells(summaryRow + 4, "D").Formula = PendingCountFormula( _
        campaignRef, approvalRef, checkedMark)

    ' Delivered totals: compare last week vs current week
    Dim deliveredCurrent As Double
    Dim deliveredLast As Double
    Dim sd As Date
    Dim delVal As Double

    deliveredCurrent = 0
    deliveredLast = 0

    For i = 1 To lo.ListRows.Count
        r = lo.DataBodyRange.Rows(i).Row
        If IsDate(ValueByHeader(wsI, r, "Send Date")) Then
            sd = DateValue(CDate(ValueByHeader(wsI, r, "Send Date")))
            If IsNumeric(ValueByHeader(wsI, r, "Delivered")) Then
                delVal = CDbl(ValueByHeader(wsI, r, "Delivered"))
            Else
                delVal = 0
            End If
            If sd >= weekStart And sd <= weekStart + 6 Then deliveredCurrent = deliveredCurrent + delVal
            If sd >= weekStart - 7 And sd <= weekStart - 1 Then deliveredLast = deliveredLast + delVal
        End If
    Next i

    wsD.Cells(summaryRow + 6, "A").Value = "Delivered Comparison"
    wsD.Cells(summaryRow + 7, "A").Value = "Last Week Delivered"
    wsD.Cells(summaryRow + 7, "B").Value = deliveredLast
    wsD.Cells(summaryRow + 8, "A").Value = "Current Week Delivered"
    wsD.Cells(summaryRow + 8, "B").Value = deliveredCurrent
End Sub

'Preserves the existing public entry point while removing fixed column letters.
Public Sub RefreshProductionStatus()
    Dim lo As ListObject
    Dim oldCalculation As XlCalculation
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    oldCalculation = Application.Calculation
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    On Error GoTo RefreshFailed

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set lo = GetInventoryTable()
    ApplyCalculatedColumns lo
    RefreshDashboard
    LogAction "RefreshProductionStatus", _
        "Production Inventory and Dashboard refreshed"

RefreshExit:
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If failureNumber <> 0 Then
        Err.Raise failureNumber, "RefreshProductionStatus", failureDescription
    End If
    Exit Sub

RefreshFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume RefreshExit
End Sub

'Public wrapper for automation scripts to apply configurations without passing ListObjects
Public Sub ApplyAllConfigurations()
    Dim lo As ListObject
    Dim checklistHeaders As Variant
    Dim item As Variant
    Dim validationResult As String
    Dim oldCalculation As XlCalculation
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    oldCalculation = Application.Calculation
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    On Error GoTo ApplyFailed
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set lo = GetInventoryTable()

    checklistHeaders = Array( _
        "Campaign Name and UTM Parameter (Source Code)", _
        "Creative Brief, SL & PH", _
        "SKUs", _
        "In-Design", _
        "Build, QA", _
        "Route", _
        "Approval", _
        "Segments")

    For Each item In checklistHeaders
        ConfigureChecklistColumn lo, CStr(item)
    Next item

    ConfigureOwnerColumn lo
    EnsureUpdatedByColumn lo
    FormatSendDateColumn lo
    ApplyCalculatedColumns lo
    RefreshDashboard
    UpdateCalendarTabs

    validationResult = ValidateWorkbookConfiguration()
    If validationResult <> "OK" Then
        Err.Raise vbObjectError + 1090, , validationResult
    End If

    LogAction "ApplyAllConfigurations", "All configurations applied successfully"

ApplyExit:
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    If failureNumber <> 0 Then
        Err.Raise failureNumber, "ApplyAllConfigurations", failureDescription
    End If
    Exit Sub

ApplyFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume ApplyExit
End Sub

'Retained for compatibility with any existing callers or UserForms.
Public Function GenerateCampaignID( _
    ByVal sendDate As Variant, _
    ByVal campaignName As String) As String

    Dim cleanName As String

    cleanName = UCase$(campaignName)
    cleanName = Replace(cleanName, "Sample | ", "")
    cleanName = Replace(cleanName, " ", "-")
    cleanName = Replace(cleanName, "/", "-")
    cleanName = Replace(cleanName, "&", "AND")
    cleanName = Replace(cleanName, "--", "-")

    If IsDate(sendDate) Then
        GenerateCampaignID = _
            "EMAIL-" & Format$(CDate(sendDate), "yyyy-mm-dd") & _
            "-" & Left$(cleanName, 25)
    Else
        GenerateCampaignID = "EMAIL-TBD-" & Left$(cleanName, 25)
    End If
End Function

'Retained with the original signature for compatibility.
Public Function CalculateCurrentStage( _
    ByVal ws As Worksheet, _
    ByVal r As Long) As String

    Dim delivered As Variant
    If Len(Trim$(CStr(ValueByHeader(ws, r, "Campaign Name")))) = 0 Then
        CalculateCurrentStage = vbNullString
        Exit Function
    End If

    delivered = ValueByHeader(ws, r, "Delivered")

    If IsNumeric(delivered) And delivered > 0 Then
        CalculateCurrentStage = "Sent"
    ElseIf Not IsChecked(ValueByHeader( _
        ws, r, "Campaign Name and UTM Parameter (Source Code)")) Then
        CalculateCurrentStage = "Source Code"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "Creative Brief, SL & PH")) Then
        CalculateCurrentStage = "Creative Brief"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "SKUs")) Then
        CalculateCurrentStage = "Waiting for SKUs"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "In-Design")) Then
        CalculateCurrentStage = "With Design"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "Build, QA")) Then
        CalculateCurrentStage = "Build / QA"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "Route")) Then
        CalculateCurrentStage = "Routing"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "Approval")) Then
        CalculateCurrentStage = "Awaiting Approval"
    ElseIf Not IsChecked(ValueByHeader(ws, r, "Segments")) Then
        CalculateCurrentStage = "Segments"
    ElseIf Not HasText(ValueByHeader(ws, r, "Jira Link")) Or _
        Not HasText(ValueByHeader(ws, r, "ClickUp Link")) Or _
        Not HasText(ValueByHeader(ws, r, "Bluecore Link")) Then
        CalculateCurrentStage = "Links Pending"
    ElseIf Not HasText(ValueByHeader(ws, r, "Est. Audience")) Then
        CalculateCurrentStage = "Ready to Schedule"
    Else
        CalculateCurrentStage = "Scheduled"
    End If
End Function

'Retained with the original signature for compatibility. Risk is now derived,
'because the stored Risk Level column is intentionally removed.
Public Function CalculateRiskLevel( _
    ByVal ws As Worksheet, _
    ByVal r As Long) As String

    Dim sendDate As Variant
    Dim currentStage As String
    Dim daysLeft As Long

    sendDate = ValueByHeader(ws, r, "Send Date")
    currentStage = CalculateCurrentStage(ws, r)

    If currentStage = "Sent" Or currentStage = "Cancelled" Then
        CalculateRiskLevel = "Complete"
        Exit Function
    End If

    If currentStage = "Scheduled" Or currentStage = "Ready to Schedule" Then
        CalculateRiskLevel = "Ready"
        Exit Function
    End If

    If Not IsDate(sendDate) Then
        CalculateRiskLevel = "At Risk"
        Exit Function
    End If

    daysLeft = DateDiff("d", Date, CDate(sendDate))

    If daysLeft <= 0 Then
        CalculateRiskLevel = "Critical"
    ElseIf daysLeft = 1 Then
        CalculateRiskLevel = "High Risk"
    ElseIf daysLeft <= 2 Then
        CalculateRiskLevel = "At Risk"
    Else
        CalculateRiskLevel = "On Track"
    End If
End Function

Public Sub RefreshDashboard()
    Dim wsD As Worksheet
    Dim wsI As Worksheet
    Dim lo As ListObject
    Dim dashboardTable As ListObject
    Dim i As Long
    Dim r As Long
    Dim matchCount As Long
    Dim displayRows As Long
    Dim displayIndex As Long
    Dim oldLastDashboardRow As Long
    Dim lastDashboardRow As Long
    Dim summaryRow As Long
    Dim dashboardValues() As Variant
    Dim jiraLinks() As String
    Dim clickUpLinks() As String
    Dim bluecoreLinks() As String
    Dim weekStart As Date
    Dim weekEnd As Date
    Dim campaignName As String
    Dim sendDate As Variant
    Dim checkedMark As String
    Dim campaignRef As String
    Dim stageRef As String
    Dim sendDateRef As String
    Dim sourceCodeRef As String
    Dim briefRef As String
    Dim buildQaRef As String
    Dim approvalRef As String

    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set wsI = ThisWorkbook.Worksheets(SH_INVENTORY)
    Set lo = GetInventoryTable()
    Set dashboardTable = wsD.ListObjects(TBL_DASHBOARD)
    checkedMark = CheckedSymbol()

    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1070, , _
            "Production Inventory has no data rows."
    End If

    weekStart = Date - Weekday(Date, vbMonday) + 1
    weekEnd = weekStart + 13

    For i = 1 To lo.ListRows.Count
        r = lo.DataBodyRange.Rows(i).Row
        campaignName = Trim$(TextValue( _
            ValueByHeader(wsI, r, "Campaign Name")))
        sendDate = ValueByHeader(wsI, r, "Send Date")

        If Len(campaignName) > 0 And Not IsError(sendDate) Then
            If IsDate(sendDate) Then
                If DateValue(CDate(sendDate)) >= weekStart And _
                    DateValue(CDate(sendDate)) <= weekEnd Then
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next i

    displayRows = matchCount
    If displayRows = 0 Then displayRows = 1

    oldLastDashboardRow = dashboardTable.Range.Row + _
        dashboardTable.Range.Rows.Count - 1
    If Not dashboardTable.DataBodyRange Is Nothing Then
        On Error Resume Next
        dashboardTable.DataBodyRange.Hyperlinks.Delete
        On Error GoTo 0
        dashboardTable.DataBodyRange.ClearContents
    End If

    wsD.Range( _
        "A" & (oldLastDashboardRow + 5) & _
        ":D" & (oldLastDashboardRow + 9)).ClearContents
    wsD.Range("A115:D119").ClearContents

    lastDashboardRow = FIRST_DASHBOARD_ROW + displayRows - 1
    summaryRow = lastDashboardRow + 5
    dashboardTable.Resize wsD.Range("A10:L" & lastDashboardRow)

    ReDim dashboardValues(1 To displayRows, 1 To 12)
    ReDim jiraLinks(1 To displayRows)
    ReDim clickUpLinks(1 To displayRows)
    ReDim bluecoreLinks(1 To displayRows)

    campaignRef = FullColumnReference(lo, "Campaign Name")
    stageRef = FullColumnReference(lo, "Current Stage")
    sendDateRef = FullColumnReference(lo, "Send Date")
    sourceCodeRef = FullColumnReference( _
        lo, "Campaign Name and UTM Parameter (Source Code)")
    briefRef = FullColumnReference(lo, "Creative Brief, SL & PH")
    buildQaRef = FullColumnReference(lo, "Build, QA")
    approvalRef = FullColumnReference(lo, "Approval")

    wsD.Range("A2").Value = _
        "Daily production view using the new checklist workflow."

    wsD.Range("A4").Value = "Active Work"
    wsD.Range("C4").Value = "Sending Today"
    wsD.Range("E4").Value = "Source Pending"
    wsD.Range("G4").Value = "Brief Pending"
    wsD.Range("I4").Value = "Approval Pending"
    wsD.Range("K4").Value = "Sent"

    wsD.Range("A6").Value = "Not sent or cancelled"
    wsD.Range("C6").Value = "Scheduled for today"
    wsD.Range("E6").Value = "Source code incomplete"
    wsD.Range("G6").Value = "Creative brief incomplete"
    wsD.Range("I6").Value = "Approval incomplete"
    wsD.Range("K6").Value = "Delivered or sent"

    wsD.Range("B3").Formula = "=NOW()"
    wsD.Range("B3").NumberFormat = "MM/DD/YYYY HH:MM"
    If wsD.Columns("B").ColumnWidth < 20 Then
        wsD.Columns("B").ColumnWidth = 20
    End If
    wsD.Range("A5").Formula = _
        "=COUNTIFS(" & campaignRef & ",""""," & _
        stageRef & ",""<>Sent""," & stageRef & ",""<>Cancelled"")"
    wsD.Range("C5").Formula = _
        "=COUNTIFS(" & sendDateRef & ",TODAY()," & campaignRef & ","""")"
    wsD.Range("E5").Formula = PendingCountFormula( _
        campaignRef, sourceCodeRef, checkedMark)
    wsD.Range("G5").Formula = PendingCountFormula( _
        campaignRef, briefRef, checkedMark)
    wsD.Range("I5").Formula = PendingCountFormula( _
        campaignRef, approvalRef, checkedMark)
    wsD.Range("K5").Formula = _
        "=COUNTIFS(" & stageRef & ",""Sent"")"

    wsD.Range("A9").Value = _
        "Rows mirror Production Inventory; checklist cells show completion."

    wsD.Range("A10").Value = "Send Date"
    wsD.Range("B10").Value = "Time"
    wsD.Range("C10").Value = "Campaign"
    wsD.Range("D10").Value = "Type"
    wsD.Range("E10").Value = "Stage"
    wsD.Range("F10").Value = "Owner"
    wsD.Range("G10").Value = "SKUs"
    wsD.Range("H10").Value = "Build, QA"
    wsD.Range("I10").Value = "Approval"
    wsD.Range("J10").Value = "Jira"
    wsD.Range("K10").Value = "ClickUp"
    wsD.Range("L10").Value = "Bluecore"

    For i = 1 To lo.ListRows.Count
        r = lo.DataBodyRange.Rows(i).Row
        campaignName = Trim$(TextValue( _
            ValueByHeader(wsI, r, "Campaign Name")))
        sendDate = ValueByHeader(wsI, r, "Send Date")

        If Len(campaignName) > 0 And Not IsError(sendDate) Then
            If IsDate(sendDate) Then
                If DateValue(CDate(sendDate)) >= weekStart And _
                    DateValue(CDate(sendDate)) <= weekEnd Then
                    displayIndex = displayIndex + 1
                    dashboardValues(displayIndex, 1) = _
                        DateValue(CDate(sendDate))
                    dashboardValues(displayIndex, 2) = _
                        ValueByHeader(wsI, r, "Send Time")
                    dashboardValues(displayIndex, 3) = campaignName
                    dashboardValues(displayIndex, 4) = _
                        ValueByHeader(wsI, r, "Campaign Type")
                    dashboardValues(displayIndex, 5) = _
                        CalculateCurrentStage(wsI, r)
                    dashboardValues(displayIndex, 6) = _
                        ValueByHeader(wsI, r, "Owner")
                    dashboardValues(displayIndex, 7) = _
                        ValueByHeader(wsI, r, "SKUs")
                    dashboardValues(displayIndex, 8) = _
                        ValueByHeader(wsI, r, "Build, QA")
                    dashboardValues(displayIndex, 9) = _
                        ValueByHeader(wsI, r, "Approval")
                    jiraLinks(displayIndex) = TextValue( _
                        ValueByHeader(wsI, r, "Jira Link"))
                    clickUpLinks(displayIndex) = TextValue( _
                        ValueByHeader(wsI, r, "ClickUp Link"))
                    bluecoreLinks(displayIndex) = TextValue( _
                        ValueByHeader(wsI, r, "Bluecore Link"))
                End If
            End If
        End If
    Next i

    dashboardTable.DataBodyRange.Value2 = dashboardValues

    For i = 1 To matchCount
        AddDashboardLink wsD.Cells(FIRST_DASHBOARD_ROW + i - 1, "J"), _
            jiraLinks(i), "Jira"
        AddDashboardLink wsD.Cells(FIRST_DASHBOARD_ROW + i - 1, "K"), _
            clickUpLinks(i), "ClickUp"
        AddDashboardLink wsD.Cells(FIRST_DASHBOARD_ROW + i - 1, "L"), _
            bluecoreLinks(i), "Bluecore"
    Next i

    wsD.Range("A11:A" & lastDashboardRow).NumberFormat = "MM/DD/YYYY"
    wsD.Range("B11:B" & lastDashboardRow).NumberFormat = "h:mm AM/PM"
    wsD.Range("G11:I" & lastDashboardRow).Font.Name = "Segoe UI Symbol"
    wsD.Range("G11:I" & lastDashboardRow).HorizontalAlignment = xlCenter

    wsD.Cells(summaryRow, "A").Value = "Today's Sends"
    wsD.Cells(summaryRow, "B").Value = "Send Date = today"
    wsD.Cells(summaryRow, "C").Value = "Confirm scheduled/send status"
    wsD.Cells(summaryRow, "D").Formula = _
        "=COUNTIFS(" & sendDateRef & ",TODAY()," & campaignRef & ","""")"

    wsD.Cells(summaryRow + 1, "A").Value = "Source Code Pending"
    wsD.Cells(summaryRow + 1, "B").Value = "Source-code checkbox is open"
    wsD.Cells(summaryRow + 1, "C").Value = _
        "Finish campaign name and UTM parameters"
    wsD.Cells(summaryRow + 1, "D").Formula = PendingCountFormula( _
        campaignRef, sourceCodeRef, checkedMark)

    wsD.Cells(summaryRow + 2, "A").Value = "Creative Brief Pending"
    wsD.Cells(summaryRow + 2, "B").Value = "Creative checkbox is open"
    wsD.Cells(summaryRow + 2, "C").Value = _
        "Complete brief, subject line, and preheader"
    wsD.Cells(summaryRow + 2, "D").Formula = PendingCountFormula( _
        campaignRef, briefRef, checkedMark)

    wsD.Cells(summaryRow + 3, "A").Value = "Build / QA Pending"
    wsD.Cells(summaryRow + 3, "B").Value = "Build/QA checkbox is open"
    wsD.Cells(summaryRow + 3, "C").Value = "Complete build and QA"
    wsD.Cells(summaryRow + 3, "D").Formula = PendingCountFormula( _
        campaignRef, buildQaRef, checkedMark)

    wsD.Cells(summaryRow + 4, "A").Value = "Approval Pending"
    wsD.Cells(summaryRow + 4, "B").Value = "Approval checkbox is open"
    wsD.Cells(summaryRow + 4, "C").Value = "Obtain final approval"
    wsD.Cells(summaryRow + 4, "D").Formula = PendingCountFormula( _
        campaignRef, approvalRef, checkedMark)
End Sub

Private Sub AddDashboardLink( _
    ByVal targetCell As Range, _
    ByVal linkAddress As String, _
    ByVal displayText As String)

    linkAddress = Trim$(linkAddress)
    If Len(linkAddress) = 0 Then Exit Sub

    targetCell.Value = displayText

    On Error Resume Next
    targetCell.Parent.Hyperlinks.Add _
        Anchor:=targetCell, _
        Address:=linkAddress, _
        TextToDisplay:=displayText
    On Error GoTo 0
End Sub

Public Sub CreateDailyDigest()
    Dim ws As Worksheet
    Dim wsD As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim r As Long
    Dim digest As String
    Dim stage As String
    Dim risk As String
    Dim campaignName As String
    Dim sendDate As Variant

    Set ws = ThisWorkbook.Worksheets(SH_INVENTORY)
    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set lo = GetInventoryTable()

    digest = "EMAIL PRODUCTION DIGEST - " & _
        Format$(Date, "mmmm d, yyyy") & vbCrLf & vbCrLf

    digest = digest & "TODAY'S SENDS:" & vbCrLf
    For i = 1 To lo.ListRows.Count
        r = lo.DataBodyRange.Rows(i).Row
        campaignName = CStr(ValueByHeader(ws, r, "Campaign Name"))
        sendDate = ValueByHeader(ws, r, "Send Date")

        If Len(Trim$(campaignName)) > 0 And IsDate(sendDate) Then
            If DateValue(CDate(sendDate)) = Date Then
                digest = digest & "- " & campaignName & " | " & _
                    CStr(ValueByHeader(ws, r, "Send Time")) & " | " & _
                    CalculateCurrentStage(ws, r) & vbCrLf
            End If
        End If
    Next i

    digest = digest & vbCrLf & "DUE SOON / OVERDUE:" & vbCrLf
    For i = 1 To lo.ListRows.Count
        r = lo.DataBodyRange.Rows(i).Row
        campaignName = CStr(ValueByHeader(ws, r, "Campaign Name"))

        If Len(Trim$(campaignName)) > 0 Then
            risk = CalculateRiskLevel(ws, r)
            If risk = "At Risk" Or risk = "High Risk" Or risk = "Critical" Then
                digest = digest & "- " & campaignName & " | " & _
                    risk & " | Stage: " & CalculateCurrentStage(ws, r) & vbCrLf
            End If
        End If
    Next i

    digest = digest & vbCrLf & "CHECKLIST PENDING:" & vbCrLf
    For i = 1 To lo.ListRows.Count
        r = lo.DataBodyRange.Rows(i).Row
        campaignName = CStr(ValueByHeader(ws, r, "Campaign Name"))
        stage = CalculateCurrentStage(ws, r)

        If Len(Trim$(campaignName)) > 0 Then
            If stage <> "Sent" And stage <> "Cancelled" And _
                stage <> "Scheduled" And stage <> "Ready to Schedule" Then
                digest = digest & "- " & campaignName & _
                    " | Next stage: " & stage & vbCrLf
            End If
        End If
    Next i

    wsD.Range("N3").Value = "Daily Digest"
    wsD.Range("N4").Value = digest
    wsD.Range("N4").WrapText = True
    wsD.Columns("N:N").ColumnWidth = 90

    LogAction "CreateDailyDigest", "Daily digest created"
    MsgBox "Daily digest created on Dashboard, starting at N3.", vbInformation
End Sub

Public Sub LogAction(ByVal actionName As String, ByVal details As String)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim logRow As ListRow
    Dim nextRow As Long

    On Error GoTo LogExit
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    Set lo = ws.ListObjects("AutomationLogTable")

    If Not lo Is Nothing Then
        Set logRow = lo.ListRows.Add
        logRow.Range.Cells(1, 1).Value = Now
        logRow.Range.Cells(1, 2).Value = CurrentUserName()
        logRow.Range.Cells(1, 3).Value = actionName
        logRow.Range.Cells(1, 4).Value = details
    Else
        nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        ws.Cells(nextRow, "A").Value = Now
        ws.Cells(nextRow, "B").Value = CurrentUserName()
        ws.Cells(nextRow, "C").Value = actionName
        ws.Cells(nextRow, "D").Value = details
    End If

LogExit:
End Sub

'Use this helper from future UserForms instead of hard-coded column letters.
Public Function InventoryColumnNumber(ByVal headerName As String) As Long
    Dim lo As ListObject
    Dim lc As ListColumn

    Set lo = GetInventoryTable()
    Set lc = FindTableColumn(lo, headerName)

    If lc Is Nothing Then
        Err.Raise vbObjectError + 1010, , _
            "Production Inventory column not found: " & headerName
    End If

    InventoryColumnNumber = lo.Range.Column + lc.Index - 1
End Function

Private Sub AddChecklistColumns( _
    ByVal lo As ListObject, _
    ByVal checklistHeaders As Variant, _
    ByVal insertAfterHeader As String)

    Dim anchor As ListColumn
    Dim lc As ListColumn
    Dim item As Variant
    Dim insertPosition As Long

    Set anchor = FindTableColumn(lo, insertAfterHeader)
    If anchor Is Nothing Then
        Err.Raise vbObjectError + 1020, , _
            "Cannot insert checklist columns after missing header: " & _
            insertAfterHeader
    End If

    insertPosition = anchor.Index + 1

    For Each item In checklistHeaders
        Set lc = FindTableColumn(lo, CStr(item))

        If lc Is Nothing Then
            Set lc = lo.ListColumns.Add(Position:=insertPosition)
            lc.Name = CStr(item)
        End If

        insertPosition = lc.Index + 1
    Next item
End Sub

Private Sub ConfigureChecklistColumn( _
    ByVal lo As ListObject, _
    ByVal headerName As String)

    Dim lc As ListColumn
    Dim rng As Range
    Dim cell As Range
    Dim listSeparator As String
    Dim uncheckedMark As String
    Dim checkedMark As String
    Dim condition As FormatCondition

    Set lc = FindTableColumn(lo, headerName)
    If lc Is Nothing Then
        Err.Raise vbObjectError + 1021, , _
            "Checklist column was not created: " & headerName
    End If

    If lc.DataBodyRange Is Nothing Then Exit Sub

    Set rng = lc.DataBodyRange
    uncheckedMark = UncheckedSymbol()
    checkedMark = CheckedSymbol()

    listSeparator = Application.International(xlListSeparator)

    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:=uncheckedMark & listSeparator & checkedMark
        .IgnoreBlank = False
        .InCellDropdown = False
        .InputTitle = "Production checklist"
        .InputMessage = "Click to toggle the checkbox status."
        .ErrorTitle = "Invalid checklist value"
        .ErrorMessage = "Choose the checked or unchecked status."
        .ShowInput = True
        .ShowError = True
    End With

    ' Initialize empty cells with unchecked symbol and double-click styling.
    For Each cell In rng.Cells
        If Len(CStr(cell.Value2)) = 0 Then
            cell.Value2 = uncheckedMark
        End If
    Next cell

    rng.NumberFormat = "@"
    rng.Font.Name = "Segoe UI Symbol"
    rng.Font.Size = 12
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter

    rng.FormatConditions.Delete

    Set condition = rng.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="=""" & checkedMark & """")
    condition.Font.Color = RGB(0, 97, 0)
    condition.Interior.Color = RGB(198, 239, 206)

    Set condition = rng.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="=""" & uncheckedMark & """")
    condition.Font.Color = RGB(89, 89, 89)
    condition.Interior.Color = RGB(242, 242, 242)

    lc.Range.ColumnWidth = 16
    lc.Range.WrapText = True
End Sub

Private Sub ConfigureOwnerColumn(ByVal lo As ListObject)
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = FindTableColumn(lo, "Owner")
    On Error GoTo 0
    If lc Is Nothing Then Exit Sub

    If Not lc.DataBodyRange Is Nothing Then
        On Error Resume Next
        lc.DataBodyRange.Validation.Delete
        On Error GoTo 0
        lc.DataBodyRange.NumberFormat = "@"
        lc.DataBodyRange.HorizontalAlignment = xlLeft
    End If
    lc.Range.ColumnWidth = 16
End Sub

Private Sub EnsureUpdatedByColumn(ByVal lo As ListObject)
    Dim updatedCol As ListColumn
    Dim createdCol As ListColumn

    Set updatedCol = FindTableColumn(lo, "Last Updated")
    If updatedCol Is Nothing Then Exit Sub

    Set createdCol = FindTableColumn(lo, "Last Updated By")
    If createdCol Is Nothing Then
        Set createdCol = lo.ListColumns.Add(Position:=updatedCol.Index + 1)
        createdCol.Name = "Last Updated By"
    End If

    If Not createdCol.DataBodyRange Is Nothing Then
        createdCol.Range.ColumnWidth = 18
        createdCol.DataBodyRange.NumberFormat = "@"
    End If
End Sub

Public Sub HandleInventorySelection(ByVal Target As Range)
    Call ToggleInventoryChecklist(Target)
End Sub

Public Function ToggleInventoryChecklist(ByVal Target As Range) As Boolean
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim columnKey As String
    Dim uncheckedMark As String
    Dim checkedMark As String

    On Error GoTo ToggleExit
    Set lo = GetInventoryTable()
    If lo Is Nothing Then Exit Function

    If Target.CountLarge <> 1 Then Exit Function
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Function

    uncheckedMark = UncheckedSymbol()
    checkedMark = CheckedSymbol()

    ' Only toggle for checklist columns
    For Each lc In lo.ListColumns
        columnKey = HeaderKey(lc.Name)
        Select Case columnKey
            Case HeaderKey("Campaign Name and UTM Parameter (Source Code)"), _
                 HeaderKey("Creative Brief, SL & PH"), _
                 HeaderKey("SKUs"), _
                 HeaderKey("In-Design"), _
                 HeaderKey("Build, QA"), _
                 HeaderKey("Route"), _
                 HeaderKey("Approval"), _
                 HeaderKey("Segments")
                If Target.Column = lo.Range.Column + lc.Index - 1 Then
                    If CStr(Target.Value) = checkedMark Then
                        Target.Value = uncheckedMark
                    Else
                        Target.Value = checkedMark
                    End If
                    UpdateRowTimestampAndUser Target.Row, lo
                    RefreshDashboard
                    ToggleInventoryChecklist = True
                    Exit Function
                End If
        End Select
    Next lc

ToggleExit:
End Function

Public Sub HandleInventoryChange(ByVal Target As Range)
    Dim lo As ListObject
    Dim changedCells As Range
    Dim rowsToStamp As Object
    Dim cell As Range
    Dim rowKey As Variant

    On Error GoTo ChangeExit

    Set lo = GetInventoryTable()
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set changedCells = Intersect(Target, lo.DataBodyRange)
    If changedCells Is Nothing Then Exit Sub

    Set rowsToStamp = CreateObject("Scripting.Dictionary")
    For Each cell In changedCells.Cells
        If IsUserEditableColumn(lo, cell.Column) Then
            rowsToStamp(CStr(cell.Row)) = cell.Row
        End If
    Next cell

    For Each rowKey In rowsToStamp.Keys
        UpdateRowTimestampAndUser CLng(rowsToStamp(rowKey)), lo
    Next rowKey

    If rowsToStamp.Count > 0 Then RefreshDashboard

ChangeExit:
End Sub

Public Sub UpdateRowTimestampAndUser(ByVal rowNumber As Long, ByVal lo As ListObject)
    Dim ws As Worksheet
    Dim tsCol As ListColumn
    Dim userCol As ListColumn

    Set ws = lo.Parent
    On Error Resume Next
    Set tsCol = FindTableColumn(lo, "Last Updated")
    Set userCol = FindTableColumn(lo, "Last Updated By")
    On Error GoTo 0

    If Not tsCol Is Nothing Then
        ws.Cells(rowNumber, tsCol.Index + lo.Range.Column - 1).Value = Now
        ws.Cells(rowNumber, tsCol.Index + lo.Range.Column - 1).NumberFormat = "MM/DD/YYYY HH:MM"
    End If

    If Not userCol Is Nothing Then
        ws.Cells(rowNumber, userCol.Index + lo.Range.Column - 1).Value = _
            CurrentUserName()
    End If
End Sub

Private Function IsUserEditableColumn( _
    ByVal lo As ListObject, _
    ByVal absoluteColumn As Long) As Boolean

    Dim lc As ListColumn
    Dim headerName As String

    For Each lc In lo.ListColumns
        If lo.Range.Column + lc.Index - 1 = absoluteColumn Then
            headerName = lc.Name
            Exit For
        End If
    Next lc

    Select Case HeaderKey(headerName)
        Case HeaderKey("Current Stage"), _
             HeaderKey("Last Updated"), _
             HeaderKey("Last Updated By")
            IsUserEditableColumn = False
        Case Else
            IsUserEditableColumn = (Len(headerName) > 0)
    End Select
End Function

Private Function CurrentUserName() As String
    CurrentUserName = Trim$(Environ$("Username"))
    If Len(CurrentUserName) = 0 Then
        CurrentUserName = Trim$(Application.UserName)
    End If
    If Len(CurrentUserName) = 0 Then
        CurrentUserName = "Unknown user"
    End If
End Function

Public Sub UpdateCalendarTabs()
    Dim monthNumber As Long
    Dim ws As Worksheet

    For monthNumber = 1 To 12
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(MonthName(monthNumber) & " Calendar")
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextMonth

        ' default tab color
        ws.Tab.Color = RGB(91, 155, 213)

        If monthNumber = Month(Date) Then
            ws.Tab.Color = RGB(0, 176, 80) ' green for current month
        End If
NextMonth:
        Set ws = Nothing
    Next monthNumber
End Sub

Private Sub FormatSendDateColumn(ByVal lo As ListObject)
    Dim lc As ListColumn

    Set lc = FindTableColumn(lo, "Send Date")
    If lc Is Nothing Then
        Err.Raise vbObjectError + 1030, , "Send Date column not found."
    End If

    If Not lc.DataBodyRange Is Nothing Then
        lc.DataBodyRange.NumberFormat = "MM/DD/YYYY"
    End If
End Sub

Private Sub ApplyCalculatedColumns(ByVal lo As ListObject)
    Dim stageColumn As ListColumn
    Dim updatedColumn As ListColumn
    Dim formulaText As String
    Dim checkedMark As String

    checkedMark = CheckedSymbol()

    Set stageColumn = FindTableColumn(lo, "Current Stage")
    Set updatedColumn = FindTableColumn(lo, "Last Updated")

    If stageColumn Is Nothing Then
        Err.Raise vbObjectError + 1031, , "Current Stage column not found."
    End If

    If updatedColumn Is Nothing Then
        Err.Raise vbObjectError + 1032, , "Last Updated column not found."
    End If

    formulaText = "=IF([@[Campaign Name]]="""","""," 
... (file truncated)
