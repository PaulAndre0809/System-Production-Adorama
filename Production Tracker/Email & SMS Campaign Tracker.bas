Attribute VB_Name = "modEmailProductionTracker"
Option Explicit

Public Const SH_EMAIL As String = "Email Campaigns"
Public Const SH_SMS As String = "SMS Campaigns"
Public Const SH_INVENTORY As String = "Email Campaigns"
Public Const SH_DASHBOARD As String = "Dashboard"
Public Const SH_LOG As String = "Automation Log"

Private Const TBL_INVENTORY As String = "EmailCampaignsTable"
Private Const TBL_SMS As String = "SMSCampaignsTable"
Private Const TBL_DASHBOARD As String = "DashboardWorkTable"
Private Const FIRST_DASHBOARD_ROW As Long = 11
Private Const CALENDAR_INVENTORY_SHEET As String = "Email Campaigns"
Private Const CALENDAR_DASHBOARD_SHEET As String = "Dashboard"
Private Const CALENDAR_TABLE As String = "EmailCampaignsTable"
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
    Dim smsTable As ListObject
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

    EnsureCampaignSheets
    Set lo = GetInventoryTable()
    Set smsTable = GetSmsTable()
    ApplyCalculatedColumns lo
    ApplyCalculatedColumns smsTable
    RefreshDashboard
    LogAction "RefreshProductionStatus", _
        "Email, SMS, calendars, and Dashboard refreshed"

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
    Dim smsTable As ListObject
    Dim checklistHeaders As Variant
    Dim smsChecklistHeaders As Variant
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
    ThisWorkbook.ForceFullCalculation = False

    EnsureCampaignSheets
    Set lo = GetInventoryTable()
    Set smsTable = GetSmsTable()

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

    smsChecklistHeaders = Array( _
        "Send SMS Options", _
        "Send Test", _
        "Approval", _
        "Segments")

    For Each item In smsChecklistHeaders
        ConfigureChecklistColumn smsTable, CStr(item)
    Next item

    ConfigureOwnerColumn lo
    EnsureUpdatedByColumn lo
    FormatSendDateColumn lo
    ApplyCalculatedColumns lo

    ConfigureOwnerColumn smsTable
    EnsureUpdatedByColumn smsTable
    FormatSendDateColumn smsTable
    ApplyCalculatedColumns smsTable

    RebuildMonthlyCalendars
    RefreshDashboard
    UpdateCalendarTabs
    StyleCoreWorkbookSheets ThisWorkbook

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
    ElseIf ws.Name = SH_SMS Then
        If Not IsChecked(ValueByHeader(ws, r, "Send SMS Options")) Then
            CalculateCurrentStage = "SMS Options"
        ElseIf Not IsChecked(ValueByHeader(ws, r, "Send Test")) Then
            CalculateCurrentStage = "Send Test"
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
    Dim emailTable As ListObject
    Dim smsTable As ListObject
    Dim dashboardTable As ListObject
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
    Dim emailCampaignRef As String
    Dim emailStageRef As String
    Dim emailDateRef As String
    Dim emailApprovalRef As String
    Dim smsCampaignRef As String
    Dim smsStageRef As String
    Dim smsDateRef As String
    Dim smsApprovalRef As String
    Dim i As Long

    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set emailTable = GetInventoryTable()
    Set smsTable = GetSmsTable()
    Set dashboardTable = wsD.ListObjects(TBL_DASHBOARD)

    weekStart = Date - Weekday(Date, vbMonday) + 1
    weekEnd = weekStart + 13

    matchCount = CountDashboardMatches(emailTable, weekStart, weekEnd) + _
        CountDashboardMatches(smsTable, weekStart, weekEnd)

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
        "A" & (oldLastDashboardRow + 1) & _
        ":D" & (oldLastDashboardRow + 10)).ClearContents
    wsD.Range("A115:D119").ClearContents

    lastDashboardRow = FIRST_DASHBOARD_ROW + displayRows - 1
    summaryRow = lastDashboardRow + 2
    dashboardTable.Resize wsD.Range("A10:L" & lastDashboardRow)

    ReDim dashboardValues(1 To displayRows, 1 To 12)
    ReDim jiraLinks(1 To displayRows)
    ReDim clickUpLinks(1 To displayRows)
    ReDim bluecoreLinks(1 To displayRows)

    emailCampaignRef = FullColumnReference(emailTable, "Campaign Name")
    emailStageRef = FullColumnReference(emailTable, "Current Stage")
    emailDateRef = FullColumnReference(emailTable, "Send Date")
    emailApprovalRef = FullColumnReference(emailTable, "Approval")
    smsCampaignRef = FullColumnReference(smsTable, "Campaign Name")
    smsStageRef = FullColumnReference(smsTable, "Current Stage")
    smsDateRef = FullColumnReference(smsTable, "Send Date")
    smsApprovalRef = FullColumnReference(smsTable, "Approval")

    wsD.Range("A2").value = _
        "Current-week and next-week Email and SMS campaign command center."
    wsD.Range("A1").value = "Email & SMS Campaign Command Center"
    wsD.Range("A8").value = "Current and Upcoming Campaigns"

    wsD.Range("A4").value = "Active Work"
    wsD.Range("C4").value = "Sending Today"
    wsD.Range("E4").value = "Email Active"
    wsD.Range("G4").value = "SMS Active"
    wsD.Range("I4").value = "Approval Pending"
    wsD.Range("K4").value = "Sent"

    wsD.Range("A6").value = "Email + SMS not yet sent"
    wsD.Range("C6").value = "Email + SMS scheduled today"
    wsD.Range("E6").value = "Open email campaigns"
    wsD.Range("G6").value = "Open SMS campaigns"
    wsD.Range("I6").value = "Approval checkbox is open"
    wsD.Range("K6").value = "Campaigns with Delivered > 0"

    wsD.Range("B3").Formula = "=NOW()"
    wsD.Range("B3").NumberFormat = "MM/DD/YYYY HH:MM"
    If wsD.Columns("B").ColumnWidth < 20 Then
        wsD.Columns("B").ColumnWidth = 20
    End If
    wsD.Range("A5").Formula = _
        "=COUNTIFS(" & emailCampaignRef & ",""<>""," & _
        emailStageRef & ",""<>Sent"")+" & _
        "COUNTIFS(" & smsCampaignRef & ",""<>""," & _
        smsStageRef & ",""<>Sent"")"
    wsD.Range("C5").Formula = _
        "=COUNTIFS(" & emailDateRef & ",TODAY()," & _
        emailCampaignRef & ",""<>"")+" & _
        "COUNTIFS(" & smsDateRef & ",TODAY()," & _
        smsCampaignRef & ",""<>"")"
    wsD.Range("E5").Formula = _
        "=COUNTIFS(" & emailCampaignRef & ",""<>""," & _
        emailStageRef & ",""<>Sent"")"
    wsD.Range("G5").Formula = _
        "=COUNTIFS(" & smsCampaignRef & ",""<>""," & _
        smsStageRef & ",""<>Sent"")"
    wsD.Range("I5").Formula = _
        "=COUNTIFS(" & emailCampaignRef & ",""<>""," & _
        emailApprovalRef & ",FALSE)+" & _
        "COUNTIFS(" & smsCampaignRef & ",""<>""," & _
        smsApprovalRef & ",FALSE)"
    wsD.Range("K5").Formula = _
        "=COUNTIF(" & emailStageRef & ",""Sent"")+" & _
        "COUNTIF(" & smsStageRef & ",""Sent"")"

    wsD.Range("A9").value = _
        "Campaigns scheduled from this Monday through next Sunday."

    wsD.Range("A10").value = "Send Date"
    wsD.Range("B10").value = "Time"
    wsD.Range("C10").value = "Channel"
    wsD.Range("D10").value = "Campaign"
    wsD.Range("E10").value = "Type"
    wsD.Range("F10").value = "Stage"
    wsD.Range("G10").value = "Owner"
    wsD.Range("H10").value = "Approval"
    wsD.Range("I10").value = "Segments"
    wsD.Range("J10").value = "Jira"
    wsD.Range("K10").value = "ClickUp"
    wsD.Range("L10").value = "Bluecore"

    AppendDashboardRows emailTable, "Email", weekStart, weekEnd, _
        dashboardValues, jiraLinks, clickUpLinks, bluecoreLinks, displayIndex
    AppendDashboardRows smsTable, "SMS", weekStart, weekEnd, _
        dashboardValues, jiraLinks, clickUpLinks, bluecoreLinks, displayIndex

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
    wsD.Range("H11:I" & lastDashboardRow).HorizontalAlignment = xlCenter

    wsD.Cells(summaryRow, "A").value = "Today's Sends"
    wsD.Cells(summaryRow, "B").value = "Email + SMS"
    wsD.Cells(summaryRow, "C").Formula = _
        "=COUNTIFS(" & emailDateRef & ",TODAY()," & _
        emailCampaignRef & ",""<>"")+" & _
        "COUNTIFS(" & smsDateRef & ",TODAY()," & _
        smsCampaignRef & ",""<>"")"
    wsD.Cells(summaryRow, "D").value = "Confirm scheduled and send status"

    wsD.Cells(summaryRow + 1, "A").value = "Email Active"
    wsD.Cells(summaryRow + 1, "B").value = "Current production"
    wsD.Cells(summaryRow + 1, "C").Formula = wsD.Range("E5").Formula
    wsD.Cells(summaryRow + 1, "D").value = "Complete the email workflow"

    wsD.Cells(summaryRow + 2, "A").value = "SMS Active"
    wsD.Cells(summaryRow + 2, "B").value = "Current production"
    wsD.Cells(summaryRow + 2, "C").Formula = wsD.Range("G5").Formula
    wsD.Cells(summaryRow + 2, "D").value = "Complete the SMS workflow"

    wsD.Cells(summaryRow + 3, "A").value = "Approval Pending"
    wsD.Cells(summaryRow + 3, "B").value = "Email + SMS"
    wsD.Cells(summaryRow + 3, "C").Formula = wsD.Range("I5").Formula
    wsD.Cells(summaryRow + 3, "D").value = "Obtain final approval"

    wsD.Cells(summaryRow + 4, "A").value = "Sent"
    wsD.Cells(summaryRow + 4, "B").value = "Delivered > 0"
    wsD.Cells(summaryRow + 4, "C").Formula = wsD.Range("K5").Formula
    wsD.Cells(summaryRow + 4, "D").value = "Completed campaigns"

    CreateDeliveredComparison wsD, emailTable
    StyleDashboard wsD, dashboardTable, summaryRow
End Sub

Private Function CountDashboardMatches( _
    ByVal lo As ListObject, _
    ByVal weekStart As Date, _
    ByVal weekEnd As Date) As Long

    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim rowNumber As Long
    Dim campaignName As String
    Dim sendDate As Variant

    Set ws = lo.Parent
    If lo.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To lo.ListRows.Count
        rowNumber = lo.DataBodyRange.Rows(rowIndex).Row
        campaignName = Trim$(TextValue( _
            ValueByHeader(ws, rowNumber, "Campaign Name")))
        sendDate = ValueByHeader(ws, rowNumber, "Send Date")

        If Len(campaignName) > 0 And IsDate(sendDate) Then
            If DateValue(CDate(sendDate)) >= weekStart And _
                DateValue(CDate(sendDate)) <= weekEnd Then
                CountDashboardMatches = CountDashboardMatches + 1
            End If
        End If
    Next rowIndex
End Function

Private Sub AppendDashboardRows( _
    ByVal lo As ListObject, _
    ByVal channelName As String, _
    ByVal weekStart As Date, _
    ByVal weekEnd As Date, _
    ByRef dashboardValues() As Variant, _
    ByRef jiraLinks() As String, _
    ByRef clickUpLinks() As String, _
    ByRef bluecoreLinks() As String, _
    ByRef displayIndex As Long)

    Dim ws As Worksheet
    Dim rowIndex As Long
    Dim rowNumber As Long
    Dim campaignName As String
    Dim sendDate As Variant

    Set ws = lo.Parent
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For rowIndex = 1 To lo.ListRows.Count
        rowNumber = lo.DataBodyRange.Rows(rowIndex).Row
        campaignName = Trim$(TextValue( _
            ValueByHeader(ws, rowNumber, "Campaign Name")))
        sendDate = ValueByHeader(ws, rowNumber, "Send Date")

        If Len(campaignName) > 0 And IsDate(sendDate) Then
            If DateValue(CDate(sendDate)) >= weekStart And _
                DateValue(CDate(sendDate)) <= weekEnd Then
                displayIndex = displayIndex + 1
                dashboardValues(displayIndex, 1) = DateValue(CDate(sendDate))
                dashboardValues(displayIndex, 2) = _
                    ValueByHeader(ws, rowNumber, "Send Time")
                dashboardValues(displayIndex, 3) = channelName
                dashboardValues(displayIndex, 4) = campaignName
                dashboardValues(displayIndex, 5) = _
                    ValueByHeader(ws, rowNumber, "Campaign Type")
                dashboardValues(displayIndex, 6) = _
                    CalculateCurrentStage(ws, rowNumber)
                dashboardValues(displayIndex, 7) = _
                    ValueByHeader(ws, rowNumber, "Owner")
                dashboardValues(displayIndex, 8) = _
                    ValueByHeader(ws, rowNumber, "Approval")
                dashboardValues(displayIndex, 9) = _
                    ValueByHeader(ws, rowNumber, "Segments")
                jiraLinks(displayIndex) = TextValue( _
                    ValueByHeader(ws, rowNumber, "Jira Link"))
                clickUpLinks(displayIndex) = TextValue( _
                    ValueByHeader(ws, rowNumber, "ClickUp Link"))
                bluecoreLinks(displayIndex) = TextValue( _
                    ValueByHeader(ws, rowNumber, "Bluecore Link"))
            End If
        End If
    Next rowIndex
End Sub

Private Sub CreateDeliveredComparison( _
    ByVal ws As Worksheet, _
    ByVal emailTable As ListObject)

    Dim comparisonTable As ListObject

    On Error Resume Next
    Set comparisonTable = ws.ListObjects("DeliveredComparisonTable")
    On Error GoTo 0
    If Not comparisonTable Is Nothing Then
        comparisonTable.Unlist
        Set comparisonTable = Nothing
    End If

    ws.Range("N2:R6").UnMerge
    ws.Range("N2:R6").Clear
    ws.Range("N2:R2").Merge
    ws.Range("N2").value = "Delivered Email Comparison"

    ws.Range("N3").value = "Period"
    ws.Range("O3").value = "Start Date"
    ws.Range("P3").value = "End Date"
    ws.Range("Q3").value = "Delivered Emails"
    ws.Range("R3").value = "Change"
    ws.Range("N4").value = "Last Week"
    ws.Range("N5").value = "Current Week"
    ws.Range("O4").Formula = "=TODAY()-WEEKDAY(TODAY(),2)-6"
    ws.Range("P4").Formula = "=O4+6"
    ws.Range("O5").Formula = "=TODAY()-WEEKDAY(TODAY(),2)+1"
    ws.Range("P5").Formula = "=O5+6"
    ws.Range("Q4").Formula = _
        "=SUMIFS(" & emailTable.Name & "[Delivered]," & _
        emailTable.Name & "[Send Date],"">=""&O4," & _
        emailTable.Name & "[Send Date],""<""&P4+1)"
    ws.Range("Q5").Formula = _
        "=SUMIFS(" & emailTable.Name & "[Delivered]," & _
        emailTable.Name & "[Send Date],"">=""&O5," & _
        emailTable.Name & "[Send Date],""<""&P5+1)"
    ws.Range("R4").value = vbNullString
    ws.Range("R5").Formula = "=Q5-Q4"

    Set comparisonTable = ws.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=ws.Range("N3:R5"), _
        XlListObjectHasHeaders:=xlYes)
    comparisonTable.Name = "DeliveredComparisonTable"

    comparisonTable.TableStyle = "TableStyleMedium2"
    ws.Range("O4:P5").NumberFormat = "MM/DD/YYYY"
    ws.Range("Q4:R5").NumberFormat = "#,##0"
End Sub

Private Sub StyleDashboard( _
    ByVal ws As Worksheet, _
    ByVal dashboardTable As ListObject, _
    ByVal summaryRow As Long)

    dashboardTable.TableStyle = "TableStyleMedium2"
    ws.Range("A1:L1").Interior.Color = RGB(31, 78, 121)
    ws.Range("A1:L1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1:L1").Font.Bold = True
    ws.Range("A1:L1").Font.Size = 20

    With ws.Range("A4:L6")
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 225, 242)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A4:L4").Font.Bold = True
    ws.Range("A5:L5").Font.Bold = True
    ws.Range("A5:L5").Font.Size = 16

    With ws.Range("N2:R2")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("A" & summaryRow & ":D" & (summaryRow + 4))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 225, 242)
        .VerticalAlignment = xlCenter
    End With
    ws.Range("A" & summaryRow & ":D" & summaryRow).Interior.Color = _
        RGB(221, 235, 247)
    ws.Range("A" & (summaryRow + 2) & ":D" & (summaryRow + 2)) _
        .Interior.Color = RGB(242, 242, 242)
    ws.Range("A" & (summaryRow + 4) & ":D" & (summaryRow + 4)) _
        .Interior.Color = RGB(242, 242, 242)

    ws.Columns("A:R").Font.Name = "Aptos"
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 11
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 34
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 20
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H:I").ColumnWidth = 12
    ws.Columns("J:L").ColumnWidth = 11
    ws.Columns("N:R").AutoFit
End Sub

Private Sub AddDashboardLink( _
    ByVal targetCell As Range, _
    ByVal linkAddress As String, _
    ByVal displayText As String)

    linkAddress = Trim$(linkAddress)
    If Len(linkAddress) = 0 Then Exit Sub

    targetCell.value = displayText

    On Error Resume Next
    targetCell.Parent.Hyperlinks.Add _
        anchor:=targetCell, _
        address:=linkAddress, _
        TextToDisplay:=displayText
    On Error GoTo 0
End Sub

Public Sub CreateDailyDigest()
    Dim ws As Worksheet
    Dim wsD As Worksheet
    Dim smsWs As Worksheet
    Dim lo As ListObject
    Dim smsTable As ListObject
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
    Set smsWs = ThisWorkbook.Worksheets(SH_SMS)
    Set smsTable = GetSmsTable()

    digest = "EMAIL & SMS CAMPAIGN DIGEST - " & _
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

    digest = digest & vbCrLf & "SMS CAMPAIGNS:" & vbCrLf
    For i = 1 To smsTable.ListRows.Count
        r = smsTable.DataBodyRange.Rows(i).Row
        campaignName = CStr(ValueByHeader(smsWs, r, "Campaign Name"))
        If Len(Trim$(campaignName)) > 0 Then
            digest = digest & "- " & campaignName & _
                " | Stage: " & CalculateCurrentStage(smsWs, r) & vbCrLf
        End If
    Next i

    wsD.Range("T2:Z20").UnMerge
    wsD.Range("T2:Z20").Clear
    wsD.Range("T2:Z2").Merge
    wsD.Range("T2").value = "Daily Digest"
    wsD.Range("T3:Z20").Merge
    wsD.Range("T3").value = digest
    wsD.Range("T3").WrapText = True
    wsD.Range("T3").VerticalAlignment = xlTop
    wsD.Columns("T:Z").ColumnWidth = 12

    LogAction "CreateDailyDigest", "Daily digest created"
    MsgBox "Daily digest created on Dashboard, starting at T2.", vbInformation
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
        logRow.Range.Cells(1, 1).value = Now
        logRow.Range.Cells(1, 2).value = CurrentUserName()
        logRow.Range.Cells(1, 3).value = actionName
        logRow.Range.Cells(1, 4).value = details
    Else
        nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        ws.Cells(nextRow, "A").value = Now
        ws.Cells(nextRow, "B").value = CurrentUserName()
        ws.Cells(nextRow, "C").value = actionName
        ws.Cells(nextRow, "D").value = details
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
    Dim values As Variant
    Dim rowIndex As Long

    Set lc = FindTableColumn(lo, headerName)
    If lc Is Nothing Then
        Err.Raise vbObjectError + 1021, , _
            "Checklist column was not created: " & headerName
    End If

    If lc.DataBodyRange Is Nothing Then Exit Sub

    Set rng = lc.DataBodyRange
    On Error Resume Next
    rng.Validation.Delete
    rng.FormatConditions.Delete
    On Error GoTo 0

    values = rng.Value2
    If rng.Cells.CountLarge = 1 Then
        rng.Value2 = IsChecked(values)
    Else
        For rowIndex = 1 To UBound(values, 1)
            values(rowIndex, 1) = IsChecked(values(rowIndex, 1))
        Next rowIndex
        rng.Value2 = values
    End If

    ' Microsoft 365 and Excel for the web use native in-cell checkboxes.
    ' Older desktop versions retain a visual, double-click Boolean fallback.
    If ApplyNativeCheckboxControl(rng) Then
        rng.NumberFormat = "General"
    Else
        ApplyLegacyCheckboxDisplay rng
    End If
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    lc.Range.ColumnWidth = 13
    lc.Range.WrapText = True
End Sub

Private Function ApplyNativeCheckboxControl(ByVal rng As Range) As Boolean
    Dim control As Object

    On Error GoTo NativeUnavailable
    Set control = CallByName(rng, "CellControl", VbGet)
    CallByName control, "SetCheckbox", VbMethod
    ApplyNativeCheckboxControl = True
    Exit Function

NativeUnavailable:
    Err.Clear
End Function

Private Function CheckboxControlType(ByVal rng As Range) As Long
    Dim control As Object

    On Error GoTo TypeUnavailable
    Set control = CallByName(rng, "CellControl", VbGet)
    CheckboxControlType = CLng(CallByName(control, "Type", VbGet))
    Exit Function

TypeUnavailable:
    CheckboxControlType = 0
    Err.Clear
End Function

Private Sub ApplyLegacyCheckboxDisplay(ByVal rng As Range)
    Dim cell As Range

    For Each cell In rng.Cells
        cell.Value2 = IIf(IsChecked(cell.Value2), 1, 0)
    Next cell

    rng.NumberFormat = LegacyCheckboxNumberFormat()
    rng.Font.Name = "Segoe UI Symbol"
End Sub

Private Sub SetChecklistValue(ByVal targetCell As Range, ByVal isComplete As Boolean)
    If CheckboxControlType(targetCell) = 2 Then
        targetCell.value = isComplete
    Else
        targetCell.Value2 = IIf(isComplete, 1, 0)
        targetCell.NumberFormat = LegacyCheckboxNumberFormat()
        targetCell.Font.Name = "Segoe UI Symbol"
    End If
End Sub

Private Function LegacyCheckboxNumberFormat() As String
    LegacyCheckboxNumberFormat = _
        "[=1]""" & CheckedSymbol() & """;" & _
        "[=0]""" & UncheckedSymbol() & """;""" & _
        UncheckedSymbol() & """"
End Function

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
    Dim targetPosition As Long

    Set updatedCol = FindTableColumn(lo, "Last Updated")
    If updatedCol Is Nothing Then Exit Sub

    Set createdCol = FindTableColumn(lo, "Last Updated By")
    targetPosition = updatedCol.Index + 1

    If createdCol Is Nothing Then
        Set createdCol = lo.ListColumns.Add(Position:=targetPosition)
        createdCol.Name = "Last Updated By"
    ElseIf createdCol.Index <> targetPosition Then
        createdCol.Range.Cut
        lo.ListColumns(targetPosition).Range.Insert Shift:=xlToRight
        Application.CutCopyMode = False
        Set createdCol = FindTableColumn(lo, "Last Updated By")
    End If

    If Not createdCol.DataBodyRange Is Nothing Then
        createdCol.Range.ColumnWidth = 18
        createdCol.DataBodyRange.NumberFormat = "@"
    End If
End Sub

Public Sub HandleInventorySelection(ByVal Target As Range)
    ' Retained for backward compatibility. Native checkboxes handle clicks.
End Sub

Public Function ToggleInventoryChecklist(ByVal Target As Range) As Boolean
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim columnKey As String

    On Error GoTo ToggleExit
    Set lo = CampaignTableForSheet(Target.Worksheet)
    If lo Is Nothing Then Exit Function

    If Target.CountLarge <> 1 Then Exit Function
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Function

    ' Retained for callers that explicitly invoke the old toggle macro.
    For Each lc In lo.ListColumns
        columnKey = HeaderKey(lc.Name)
        Select Case columnKey
            Case HeaderKey("Campaign Name and UTM Parameter (Source Code)"), _
                 HeaderKey("Creative Brief, SL & PH"), _
                 HeaderKey("SKUs"), _
                 HeaderKey("In-Design"), _
                 HeaderKey("Build, QA"), _
                 HeaderKey("Route"), _
                 HeaderKey("Send SMS Options"), _
                 HeaderKey("Send Test"), _
                 HeaderKey("Approval"), _
                HeaderKey("Segments")
                If Target.Column = lo.Range.Column + lc.Index - 1 Then
                    SetChecklistValue Target, Not IsChecked(Target.value)
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
    HandleCampaignChange Target.Worksheet, Target
End Sub

Public Sub HandleCampaignChange( _
    ByVal ws As Worksheet, _
    ByVal Target As Range)

    Dim lo As ListObject
    Dim changedCells As Range
    Dim rowsToStamp As Object
    Dim cell As Range
    Dim rowKey As Variant

    On Error GoTo ChangeExit

    Set lo = CampaignTableForSheet(ws)
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
        CalculateStageRow CLng(rowsToStamp(rowKey)), lo
    Next rowKey

    If rowsToStamp.Count > 0 Then RefreshDashboard

ChangeExit:
End Sub

Private Sub CalculateStageRow(ByVal rowNumber As Long, ByVal lo As ListObject)
    Dim stageColumn As ListColumn
    Dim relativeRow As Long

    Set stageColumn = FindTableColumn(lo, "Current Stage")
    If stageColumn Is Nothing Then Exit Sub
    If stageColumn.DataBodyRange Is Nothing Then Exit Sub

    relativeRow = rowNumber - lo.DataBodyRange.Row + 1
    If relativeRow < 1 Or relativeRow > lo.ListRows.Count Then Exit Sub

    stageColumn.DataBodyRange.Cells(relativeRow, 1).Calculate
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
        ws.Cells(rowNumber, tsCol.Index + lo.Range.Column - 1).value = Now
        ws.Cells(rowNumber, tsCol.Index + lo.Range.Column - 1).NumberFormat = "MM/DD/YYYY HH:MM"
    End If

    If Not userCol Is Nothing Then
        ws.Cells(rowNumber, userCol.Index + lo.Range.Column - 1).value = _
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
        If ws Is Nothing Then GoTo nextMonth

        ' default tab color
        ws.Tab.Color = RGB(91, 155, 213)

        If monthNumber = Month(Date) Then
            ws.Tab.Color = RGB(0, 176, 80) ' green for current month
        End If
nextMonth:
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
    Set stageColumn = FindTableColumn(lo, "Current Stage")
    Set updatedColumn = FindTableColumn(lo, "Last Updated")

    If stageColumn Is Nothing Then
        Err.Raise vbObjectError + 1031, , "Current Stage column not found."
    End If

    If updatedColumn Is Nothing Then
        Err.Raise vbObjectError + 1032, , "Last Updated column not found."
    End If

    formulaText = "=IF([@[Campaign Name]]="""","""","
    formulaText = formulaText & _
        "IF(AND(ISNUMBER([@Delivered]),[@Delivered]>0),""Sent"","

    If lo.Parent.Name = SH_SMS Then
        formulaText = formulaText & _
            "IF(NOT([@[Send SMS Options]]),""SMS Options"","
        formulaText = formulaText & _
            "IF(NOT([@[Send Test]]),""Send Test"","
        formulaText = formulaText & _
            "IF(NOT([@Approval]),""Awaiting Approval"","
        formulaText = formulaText & _
            "IF(NOT([@Segments]),""Segments"","
        formulaText = formulaText & _
            "IF(OR(IFERROR(LEN(TRIM([@[Jira Link]]&"""")),0)=0," & _
            "IFERROR(LEN(TRIM([@[ClickUp Link]]&"""")),0)=0," & _
            "IFERROR(LEN(TRIM([@[Bluecore Link]]&"""")),0)=0),""Links Pending"","
        formulaText = formulaText & _
            "IF(IFERROR(LEN(TRIM([@[Est. Audience]]&"""")),0)=0," & _
            """Ready to Schedule"",""Scheduled""" & String$(8, ")")
    Else
        formulaText = formulaText & _
            "IF(NOT([@[Campaign Name and UTM Parameter (Source Code)]]),""Source Code"","
        formulaText = formulaText & _
            "IF(NOT([@[Creative Brief, SL & PH]]),""Creative Brief"","
        formulaText = formulaText & _
            "IF(NOT([@SKUs]),""Waiting for SKUs"","
        formulaText = formulaText & _
            "IF(NOT([@[In-Design]]),""With Design"","
        formulaText = formulaText & _
            "IF(NOT([@[Build, QA]]),""Build / QA"","
        formulaText = formulaText & _
            "IF(NOT([@Route]),""Routing"","
        formulaText = formulaText & _
            "IF(NOT([@Approval]),""Awaiting Approval"","
        formulaText = formulaText & _
            "IF(NOT([@Segments]),""Segments"","
        formulaText = formulaText & _
            "IF(OR(IFERROR(LEN(TRIM([@[Jira Link]]&"""")),0)=0," & _
            "IFERROR(LEN(TRIM([@[ClickUp Link]]&"""")),0)=0," & _
            "IFERROR(LEN(TRIM([@[Bluecore Link]]&"""")),0)=0),""Links Pending"","
        formulaText = formulaText & _
            "IF(IFERROR(LEN(TRIM([@[Est. Audience]]&"""")),0)=0," & _
            """Ready to Schedule"",""Scheduled""" & String$(12, ")")
    End If

    If Not stageColumn.DataBodyRange Is Nothing Then
        stageColumn.DataBodyRange.Formula = formulaText
        stageColumn.DataBodyRange.Calculate
    End If

    ' Convert any legacy NOW() formulas to their current values. Existing audit
    ' timestamps must survive every refresh.
    If Not updatedColumn.DataBodyRange Is Nothing Then
        updatedColumn.DataBodyRange.Value2 = updatedColumn.DataBodyRange.Value2
        FormatLastUpdatedColumn updatedColumn
    End If
End Sub

Private Sub FormatLastUpdatedColumn(ByVal updatedColumn As ListColumn)
    Const TIMESTAMP_FORMAT As String = "MM/DD/YYYY HH:MM"
    Const MINIMUM_COLUMN_WIDTH As Double = 20

    If Not updatedColumn.DataBodyRange Is Nothing Then
        updatedColumn.DataBodyRange.NumberFormat = TIMESTAMP_FORMAT
        updatedColumn.DataBodyRange.HorizontalAlignment = xlLeft
    End If

    updatedColumn.Range.EntireColumn.AutoFit

    If updatedColumn.Range.ColumnWidth < MINIMUM_COLUMN_WIDTH Then
        updatedColumn.Range.ColumnWidth = MINIMUM_COLUMN_WIDTH
    End If
End Sub

Private Sub DeleteTableColumnIfPresent( _
    ByVal lo As ListObject, _
    ByVal headerName As String)

    Dim lc As ListColumn

    Set lc = FindTableColumn(lo, headerName)
    If Not lc Is Nothing Then lc.Delete
End Sub

Private Sub RequireTableColumn( _
    ByVal lo As ListObject, _
    ByVal headerName As String)

    Dim lc As ListColumn

    Set lc = FindTableColumn(lo, headerName)

    If lc Is Nothing Then
        Err.Raise vbObjectError + 1040, , _
            "Required Production Inventory column not found: " & headerName
    End If
End Sub

Private Sub EnsureCampaignSheets()
    Dim wb As Workbook
    Dim emailSheet As Worksheet
    Dim smsSheet As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim columnIndex As Long

    Set wb = ThisWorkbook

    On Error Resume Next
    Set emailSheet = wb.Worksheets(SH_EMAIL)
    On Error GoTo 0

    If emailSheet Is Nothing Then
        On Error Resume Next
        Set emailSheet = wb.Worksheets("Production Inventory")
        On Error GoTo 0

        If emailSheet Is Nothing Then
            Err.Raise vbObjectError + 1041, , _
                "Email Campaigns sheet was not found."
        End If
        emailSheet.Name = SH_EMAIL
    End If

    On Error Resume Next
    Set lo = emailSheet.ListObjects(TBL_INVENTORY)
    If lo Is Nothing Then
        Set lo = emailSheet.ListObjects("ProductionInventoryTable")
        If Not lo Is Nothing Then lo.Name = TBL_INVENTORY
    End If
    On Error GoTo 0

    If lo Is Nothing Then
        Err.Raise vbObjectError + 1042, , _
            "Email campaign table was not found."
    End If

    On Error Resume Next
    Set smsSheet = wb.Worksheets(SH_SMS)
    On Error GoTo 0

    If smsSheet Is Nothing Then
        Set smsSheet = wb.Worksheets.Add(After:=emailSheet)
        smsSheet.Name = SH_SMS
    End If

    Set lo = Nothing
    On Error Resume Next
    Set lo = smsSheet.ListObjects(TBL_SMS)
    On Error GoTo 0

    If lo Is Nothing Then
        If Application.WorksheetFunction.CountA(smsSheet.UsedRange) > 0 Then
            Err.Raise vbObjectError + 1043, , _
                "SMS Campaigns contains data but no SMSCampaignsTable."
        End If

        headers = Array( _
            "Send Date", _
            "Send Time", _
            "Campaign Name", _
            "Campaign Type", _
            "Current Stage", _
            "Owner", _
            "Send SMS Options", _
            "Send Test", _
            "Approval", _
            "Segments", _
            "Jira Link", _
            "ClickUp Link", _
            "Bluecore Link", _
            "Est. Audience", _
            "Delivered", _
            "Last Updated", _
            "Last Updated By")

        For columnIndex = LBound(headers) To UBound(headers)
            smsSheet.Cells(1, columnIndex + 1).value = headers(columnIndex)
        Next columnIndex

        Set lo = smsSheet.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=smsSheet.Range("A1:Q201"), _
            XlListObjectHasHeaders:=xlYes)
        lo.Name = TBL_SMS
    End If

    lo.TableStyle = "TableStyleMedium2"
End Sub

Private Function GetInventoryTable() As ListObject
    Set GetInventoryTable = _
        ThisWorkbook.Worksheets(SH_INVENTORY).ListObjects(TBL_INVENTORY)
End Function

Private Function GetSmsTable() As ListObject
    Set GetSmsTable = _
        ThisWorkbook.Worksheets(SH_SMS).ListObjects(TBL_SMS)
End Function

Private Function FindTableColumn( _
    ByVal lo As ListObject, _
    ByVal headerName As String) As ListColumn

    Dim lc As ListColumn
    Dim requestedKey As String

    requestedKey = HeaderKey(headerName)

    For Each lc In lo.ListColumns
        If HeaderKey(lc.Name) = requestedKey Then
            Set FindTableColumn = lc
            Exit Function
        End If
    Next lc
End Function

Private Function HeaderKey(ByVal headerName As String) As String
    Dim normalizedValue As String

    normalizedValue = LCase$(Trim$(headerName))
    normalizedValue = Replace(normalizedValue, " ", "")
    normalizedValue = Replace(normalizedValue, "/", "")
    normalizedValue = Replace(normalizedValue, "\", "")
    normalizedValue = Replace(normalizedValue, "-", "")
    normalizedValue = Replace(normalizedValue, "_", "")
    normalizedValue = Replace(normalizedValue, ".", "")
    normalizedValue = Replace(normalizedValue, ",", "")
    normalizedValue = Replace(normalizedValue, "&", "")
    normalizedValue = Replace(normalizedValue, "(", "")
    normalizedValue = Replace(normalizedValue, ")", "")

    HeaderKey = normalizedValue
End Function

Private Function ValueByHeader( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal headerName As String) As Variant

    Dim lo As ListObject
    Dim lc As ListColumn

    Set lo = CampaignTableForSheet(ws)
    Set lc = FindTableColumn(lo, headerName)

    If lc Is Nothing Then
        Err.Raise vbObjectError + 1050, , _
            "Campaign column not found on " & ws.Name & ": " & headerName
    End If

    ValueByHeader = ws.Cells( _
        rowNumber, lo.Range.Column + lc.Index - 1).value
End Function

Private Function CampaignTableForSheet(ByVal ws As Worksheet) As ListObject
    Select Case ws.Name
        Case SH_EMAIL
            Set CampaignTableForSheet = GetInventoryTable()
        Case SH_SMS
            Set CampaignTableForSheet = GetSmsTable()
        Case Else
            Err.Raise vbObjectError + 1052, , _
                "Unsupported campaign sheet: " & ws.Name
    End Select
End Function

Private Function IsChecked(ByVal inputValue As Variant) As Boolean
    Dim TextValue As String

    If VarType(inputValue) = vbBoolean Then
        IsChecked = CBool(inputValue)
        Exit Function
    End If

    If IsNumeric(inputValue) Then
        IsChecked = (CDbl(inputValue) <> 0)
        Exit Function
    End If

    TextValue = LCase$(Trim$(CStr(inputValue)))

    IsChecked = _
        (TextValue = LCase$(CheckedSymbol())) Or _
        (TextValue = "true") Or _
        (TextValue = "yes") Or _
        (TextValue = "x") Or _
        (TextValue = "done") Or _
        (TextValue = "complete") Or _
        (TextValue = "completed")
End Function

Private Function HasText(ByVal inputValue As Variant) As Boolean
    If IsError(inputValue) Or IsNull(inputValue) Or _
        IsEmpty(inputValue) Then Exit Function

    HasText = (Len(Trim$(CStr(inputValue))) > 0)
End Function

Private Function TextValue(ByVal inputValue As Variant) As String
    If IsError(inputValue) Or IsNull(inputValue) Or _
        IsEmpty(inputValue) Then Exit Function
    TextValue = CStr(inputValue)
End Function

Private Function CheckedSymbol() As String
    CheckedSymbol = ChrW$(&H2611)
End Function

Private Function UncheckedSymbol() As String
    UncheckedSymbol = ChrW$(&H2610)
End Function

Private Function FullColumnReference( _
    ByVal lo As ListObject, _
    ByVal headerName As String) As String

    Dim lc As ListColumn

    Set lc = FindTableColumn(lo, headerName)
    If lc Is Nothing Then
        Err.Raise vbObjectError + 1051, , _
            "Production Inventory column not found: " & headerName
    End If

    FullColumnReference = "'" & lo.Parent.Name & "'!" & _
        lc.DataBodyRange.address
End Function

Private Function PendingCountFormula( _
    ByVal campaignRef As String, _
    ByVal checklistRef As String) As String

    PendingCountFormula = _
        "=COUNTIFS(" & campaignRef & ",""<>""," & _
        checklistRef & ",FALSE)"
End Function

Private Function CreateBackupCopy(ByVal wb As Workbook) As String
    Dim dotPosition As Long
    Dim baseName As String
    Dim extension As String
    Dim backupPath As String

    If Len(wb.Path) = 0 Then
        Err.Raise vbObjectError + 1060, , _
            "Save the workbook before running the migration."
    End If

    dotPosition = InStrRev(wb.Name, ".")

    If dotPosition > 0 Then
        baseName = Left$(wb.Name, dotPosition - 1)
        extension = Mid$(wb.Name, dotPosition)
    Else
        baseName = wb.Name
        extension = ".xlsm"
    End If

    backupPath = wb.Path & Application.PathSeparator & _
        baseName & "_PRE_MIGRATION_" & _
        Format$(Now, "yyyymmdd_hhnnss") & extension

    wb.SaveCopyAs backupPath
    CreateBackupCopy = backupPath
End Function

Private Function CountBrokenReferences(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim formulas As Range
    Dim cell As Range
    Dim nm As Name
    Dim total As Long

    For Each ws In wb.Worksheets
        Set formulas = Nothing

        On Error Resume Next
        Set formulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0

        If Not formulas Is Nothing Then
            For Each cell In formulas.Cells
                If InStr(1, CStr(cell.Formula), "#REF!", vbTextCompare) > 0 Then
                    total = total + 1
                End If
            Next cell
        End If
    Next ws

    For Each nm In wb.Names
        On Error Resume Next
        If InStr(1, CStr(nm.RefersTo), "#REF!", vbTextCompare) > 0 Then
            total = total + 1
        End If
        On Error GoTo 0
    Next nm

    CountBrokenReferences = total
End Function

Public Function ValidateWorkbookConfiguration() As String
    Dim emailTable As ListObject
    Dim smsTable As ListObject
    Dim dashboardTable As ListObject
    Dim comparisonTable As ListObject
    Dim emailHeaders As Variant
    Dim smsHeaders As Variant
    Dim emailChecklist As Variant
    Dim smsChecklist As Variant
    Dim monthNumber As Long
    Dim monthSheet As Worksheet
    Dim expectedColor As Long

    On Error GoTo ValidationFailed

    Set emailTable = GetInventoryTable()
    Set smsTable = GetSmsTable()
    Set dashboardTable = ThisWorkbook.Worksheets(SH_DASHBOARD) _
        .ListObjects(TBL_DASHBOARD)

    emailHeaders = Array( _
        "Send Date", _
        "Send Time", _
        "Campaign Name", _
        "Campaign Type", _
        "Current Stage", _
        "Owner", _
        "Campaign Name and UTM Parameter (Source Code)", _
        "Creative Brief, SL & PH", _
        "SKUs", _
        "In-Design", _
        "Build, QA", _
        "Route", _
        "Approval", _
        "Segments", _
        "Jira Link", _
        "ClickUp Link", _
        "Bluecore Link", _
        "Est. Audience", _
        "Delivered", _
        "Last Updated", _
        "Last Updated By")

    emailChecklist = Array( _
        "Campaign Name and UTM Parameter (Source Code)", _
        "Creative Brief, SL & PH", _
        "SKUs", _
        "In-Design", _
        "Build, QA", _
        "Route", _
        "Approval", _
        "Segments")

    smsHeaders = Array( _
        "Send Date", _
        "Send Time", _
        "Campaign Name", _
        "Campaign Type", _
        "Current Stage", _
        "Owner", _
        "Send SMS Options", _
        "Send Test", _
        "Approval", _
        "Segments", _
        "Jira Link", _
        "ClickUp Link", _
        "Bluecore Link", _
        "Est. Audience", _
        "Delivered", _
        "Last Updated", _
        "Last Updated By")

    smsChecklist = Array( _
        "Send SMS Options", _
        "Send Test", _
        "Approval", _
        "Segments")

    ValidateCampaignTable emailTable, emailHeaders, emailChecklist
    ValidateCampaignTable smsTable, smsHeaders, smsChecklist

    For monthNumber = 1 To 12
        Set monthSheet = ThisWorkbook.Worksheets( _
            MonthName(monthNumber) & " Calendar")
        If monthNumber = Month(Date) Then
            expectedColor = RGB(0, 176, 80)
        Else
            expectedColor = RGB(91, 155, 213)
        End If
        If monthSheet.Tab.Color <> expectedColor Then
            Err.Raise vbObjectError + 1087, , _
                "Calendar tab color is incorrect: " & monthSheet.Name
        End If
        If InStr(1, monthSheet.Range("A6").Formula, _
            TBL_INVENTORY, vbTextCompare) = 0 Or _
            InStr(1, monthSheet.Range("A6").Formula, _
            TBL_SMS, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1088, , _
                "Calendar does not include both campaign tables: " & _
                monthSheet.Name
        End If
    Next monthNumber

    If dashboardTable.ListColumns.Count <> 12 Or _
        dashboardTable.ListRows.Count < 1 Then
        Err.Raise vbObjectError + 1081, , _
            "Dashboard work table has an invalid structure."
    End If

    On Error Resume Next
    Set comparisonTable = ThisWorkbook.Worksheets(SH_DASHBOARD) _
        .ListObjects("DeliveredComparisonTable")
    On Error GoTo ValidationFailed
    If comparisonTable Is Nothing Or comparisonTable.ListColumns.Count <> 5 Then
        Err.Raise vbObjectError + 1089, , _
            "Delivered email comparison table is missing or invalid."
    End If

    If CountBrokenReferences(ThisWorkbook) > 0 Then
        Err.Raise vbObjectError + 1082, , _
            "The workbook contains one or more #REF! formulas or names."
    End If

    ValidateWorkbookConfiguration = "OK"
    Exit Function

ValidationFailed:
    ValidateWorkbookConfiguration = "QA failed: " & Err.Description
End Function

Private Sub ValidateCampaignTable( _
    ByVal lo As ListObject, _
    ByVal requiredHeaders As Variant, _
    ByVal checklistHeaders As Variant)

    Dim item As Variant
    Dim lc As ListColumn
    Dim ownerCol As ListColumn
    Dim updatedCol As ListColumn
    Dim updatedByCol As ListColumn
    Dim stageCol As ListColumn
    Dim controlType As Long

    For Each item In requiredHeaders
        Set lc = FindTableColumn(lo, CStr(item))
        If lc Is Nothing Then
            Err.Raise vbObjectError + 1080, , _
                "Missing " & lo.Parent.Name & " column: " & CStr(item)
        End If
    Next item

    Set ownerCol = FindTableColumn(lo, "Owner")
    If RangeHasValidation(ownerCol.DataBodyRange) Then
        Err.Raise vbObjectError + 1083, , _
            lo.Parent.Name & " Owner must be plain text."
    End If

    For Each item In checklistHeaders
        Set lc = FindTableColumn(lo, CStr(item))
        If RangeHasValidation(lc.DataBodyRange) Then
            Err.Raise vbObjectError + 1084, , _
                "Checklist column contains a dropdown: " & CStr(item)
        End If

        controlType = CheckboxControlType(lc.DataBodyRange)
        If controlType <> 2 And _
            InStr(1, lc.DataBodyRange.NumberFormat, _
            CheckedSymbol(), vbBinaryCompare) = 0 Then
            Err.Raise vbObjectError + 1085, , _
                "Checklist column is not checkbox-formatted: " & CStr(item)
        End If
    Next item

    Set updatedCol = FindTableColumn(lo, "Last Updated")
    Set updatedByCol = FindTableColumn(lo, "Last Updated By")
    If updatedByCol.Index <> updatedCol.Index + 1 Then
        Err.Raise vbObjectError + 1086, , _
            lo.Parent.Name & _
            " Last Updated By must be right of Last Updated."
    End If

    Set stageCol = FindTableColumn(lo, "Current Stage")
    If stageCol.DataBodyRange Is Nothing Or _
        Not stageCol.DataBodyRange.HasFormula Then
        Err.Raise vbObjectError + 1087, , _
            lo.Parent.Name & " Current Stage is not formula-driven."
    End If
End Sub

Private Function RangeHasValidation(ByVal rng As Range) As Boolean
    Dim validationType As Long

    If rng Is Nothing Then Exit Function

    On Error Resume Next
    validationType = rng.Validation.Type
    RangeHasValidation = (Err.Number = 0 And validationType <> 0)
    Err.Clear
    On Error GoTo 0
End Function

'Rebuilds the twelve current-year calendar sheets and workbook navigation.
Public Sub RebuildMonthlyCalendars()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim monthNumber As Long
    Dim oldCalculation As XlCalculation
    Dim oldDisplayAlerts As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldScreenUpdating As Boolean

    Set wb = ThisWorkbook

    oldCalculation = Application.Calculation
    oldDisplayAlerts = Application.DisplayAlerts
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    On Error GoTo BuildFailed

    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For monthNumber = 1 To 12
        Set ws = GetOrCreateCalendarSheet(wb, monthNumber)
        BuildCalendarSheet ws, monthNumber
    Next monthNumber

    DeleteSheetIfPresent wb, "README"
    DeleteSheetIfPresent wb, "VBA Code"

    AddDashboardCalendarLinks wb
    StyleCoreWorkbookSheets wb
    UpdateCalendarTabs
    OrderWorkbookSheets wb
    ConfigureWorkbookViews wb

    wb.Worksheets("Dropdowns").Visible = xlSheetVeryHidden
    wb.Worksheets("Automation Log").Visible = xlSheetVeryHidden

    For monthNumber = 1 To 12
        wb.Worksheets(MonthName(monthNumber) & " Calendar").Calculate
    Next monthNumber
    LogAction "RebuildMonthlyCalendars", _
        "Twelve dynamic monthly calendar sheets rebuilt"

BuildExit:
    Application.Calculation = oldCalculation
    Application.DisplayAlerts = oldDisplayAlerts
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

BuildFailed:
    MsgBox "Calendar rebuild stopped: " & Err.Description, _
        vbCritical, "Calendar Rebuild"
    Resume BuildExit
End Sub

Private Function GetOrCreateCalendarSheet( _
    ByVal wb As Workbook, _
    ByVal monthNumber As Long) As Worksheet

    Dim targetName As String
    Dim legacySheet As Worksheet

    targetName = MonthName(monthNumber) & " Calendar"

    On Error Resume Next
    Set GetOrCreateCalendarSheet = wb.Worksheets(targetName)
    On Error GoTo 0

    If Not GetOrCreateCalendarSheet Is Nothing Then Exit Function

    If monthNumber = Month(Date) Then
        On Error Resume Next
        Set legacySheet = wb.Worksheets("Calendar Import")
        On Error GoTo 0

        If Not legacySheet Is Nothing Then
            legacySheet.Name = targetName
            Set GetOrCreateCalendarSheet = legacySheet
            Exit Function
        End If
    End If

    Set GetOrCreateCalendarSheet = wb.Worksheets.Add( _
        After:=wb.Worksheets(wb.Worksheets.Count))
    GetOrCreateCalendarSheet.Name = targetName
End Function

Private Sub BuildCalendarSheet( _
    ByVal ws As Worksheet, _
    ByVal monthNumber As Long)

    Dim calendarRange As Range
    Dim dayNames As Variant
    Dim dayIndex As Long
    Dim previousMonth As Long
    Dim nextMonth As Long
    Dim formulaText As String
    Dim bullet As String
    Dim condition As FormatCondition

    ClearCalendarSheet ws

    bullet = ChrW$(&H2022)
    previousMonth = IIf(monthNumber = 1, 12, monthNumber - 1)
    nextMonth = IIf(monthNumber = 12, 1, monthNumber + 1)

    ws.Range("A1:G1").Merge
    ws.Range("A1").Formula = _
        "=TEXT(DATE(YEAR(TODAY())," & monthNumber & _
        ",1),""mmmm yyyy"")&"" Campaign Calendar"""

    ws.Range("A2:G2").Merge
    ws.Range("A2").value = _
        "Email and SMS campaigns update automatically from their send dates."

    SetInternalLink ws, ws.Range("A3"), _
        "Dashboard", "'Dashboard'!A1"
    SetInternalLink ws, ws.Range("B3"), _
        "Email", "'Email Campaigns'!A1"
    SetInternalLink ws, ws.Range("C3"), _
        "SMS", "'SMS Campaigns'!A1"

    ws.Range("D3:E3").Merge
    ws.Range("D3").Formula = _
        "=""Today: ""&TEXT(TODAY(),""MM/DD/YYYY"")"

    SetInternalLink ws, ws.Range("F3"), _
        "< " & MonthName(previousMonth), _
        "'" & MonthName(previousMonth) & " Calendar'!A1"
    SetInternalLink ws, ws.Range("G3"), _
        MonthName(nextMonth) & " >", _
        "'" & MonthName(nextMonth) & " Calendar'!A1"

    ws.Range("A4:G4").Merge
    ws.Range("A4").value = _
        "Each campaign appears on its scheduled Send Date. " & _
        "Multiple campaigns are listed together."

    dayNames = Array( _
        "Sunday", "Monday", "Tuesday", "Wednesday", _
        "Thursday", "Friday", "Saturday")

    For dayIndex = 0 To 6
        ws.Cells(5, dayIndex + 1).value = dayNames(dayIndex)
    Next dayIndex

    formulaText = _
        "=LET(first,DATE(YEAR(TODAY())," & monthNumber & ",1)," & _
        "d,first-WEEKDAY(first,1)+1+(ROW()-6)*7+COLUMN()-1," & _
        "items,TEXTJOIN(CHAR(10)&""" & bullet & " "",TRUE," & _
        "FILTER(""Email | ""&" & CALENDAR_TABLE & "[Campaign Name]," & _
        "IFERROR(INT(" & CALENDAR_TABLE & "[Send Date]),0)=d,"""")," & _
        "FILTER(""SMS | ""&" & TBL_SMS & "[Campaign Name]," & _
        "IFERROR(INT(" & TBL_SMS & "[Send Date]),0)=d,""""))," & _
        "IF(MONTH(d)<>" & monthNumber & ",""""," & _
        "DAY(d)&IF(items="""","""",CHAR(10)&""" & bullet & " ""&items)))"

    Set calendarRange = ws.Range( _
        "A" & FIRST_CALENDAR_ROW & ":G" & LAST_CALENDAR_ROW)
    SetDynamicFormula calendarRange, formulaText

    ws.Range("A12:G12").Merge
    ws.Range("A12").value = _
        "Update Send Date or Campaign Name in Email Campaigns or " & _
        "SMS Campaigns to refresh this calendar."

    StyleCalendarSheet ws, monthNumber

    calendarRange.FormatConditions.Delete

    Set condition = calendarRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=ISNUMBER(SEARCH(CHAR(10),A6))")
    condition.Interior.Color = RGB(226, 239, 218)

    Set condition = calendarRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=DATE(YEAR(TODAY())," & monthNumber & _
            ",1)-WEEKDAY(DATE(YEAR(TODAY())," & monthNumber & _
            ",1),1)+1+(ROW()-6)*7+COLUMN()-1=TODAY()")
    condition.Interior.Color = RGB(255, 235, 156)
    condition.Font.Bold = True
End Sub

Private Sub SetDynamicFormula( _
    ByVal targetRange As Range, _
    ByVal formulaText As String)

    On Error Resume Next
    CallByName targetRange, "Formula2", VbLet, formulaText
    If Err.Number <> 0 Then
        Err.Clear
        targetRange.Formula = formulaText
    End If
    On Error GoTo 0
End Sub

Private Sub ClearCalendarSheet(ByVal ws As Worksheet)
    Do While ws.ListObjects.Count > 0
        ws.ListObjects(1).Delete
    Loop

    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop

    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.FormatConditions.Delete
    ws.Hyperlinks.Delete
End Sub

Private Sub StyleCalendarSheet( _
    ByVal ws As Worksheet, _
    ByVal monthNumber As Long)

    Dim rowNumber As Long

    ws.Cells.Font.Name = "Aptos"
    ws.Cells.Font.Strikethrough = False
    ws.Columns("A:G").ColumnWidth = 22
    ws.Rows(1).RowHeight = 34
    ws.Rows(2).RowHeight = 22
    ws.Rows(3).RowHeight = 24
    ws.Rows(4).RowHeight = 28
    ws.Rows(5).RowHeight = 24

    For rowNumber = FIRST_CALENDAR_ROW To LAST_CALENDAR_ROW
        ws.Rows(rowNumber).RowHeight = 88
    Next rowNumber

    ws.Rows(12).RowHeight = 22

    With ws.Range("A1:G1")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("A2:G2")
        .Font.Color = RGB(89, 89, 89)
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("A3:G3")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Range("A3:C3").Interior.Color = RGB(221, 235, 247)
    ws.Range("D3:E3").Interior.Color = RGB(242, 242, 242)
    ws.Range("F3:G3").Interior.Color = RGB(221, 235, 247)

    With ws.Range("A4:G4")
        .Interior.Color = RGB(242, 242, 242)
        .Font.Color = RGB(89, 89, 89)
        .HorizontalAlignment = xlCenter
        .WrapText = True
    End With

    With ws.Range("A5:G5")
        .Interior.Color = RGB(91, 155, 213)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(166, 166, 166)
    End With

    With ws.Range( _
        "A" & FIRST_CALENDAR_ROW & ":G" & LAST_CALENDAR_ROW)
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .Font.Size = 10
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(166, 166, 166)
    End With

    ws.Range("A6:A11").Interior.Color = RGB(234, 243, 248)
    ws.Range("G6:G11").Interior.Color = RGB(234, 243, 248)

    With ws.Range("A12:G12")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With

    ' Default tab color for month sheets; current month is highlighted by UpdateCalendarTabs
    ws.Tab.Color = RGB(91, 155, 213)

    ws.PageSetup.PrintArea = "$A$1:$G$12"

    On Error Resume Next
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    On Error GoTo 0
End Sub

Private Sub AddDashboardCalendarLinks(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim monthNumber As Long

    Set ws = wb.Worksheets(CALENDAR_DASHBOARD_SHEET)

    ws.Range("A7:L7").Clear
    ws.Range("A7:L7").Hyperlinks.Delete

    For monthNumber = 1 To 12
        SetInternalLink ws, ws.Cells(7, monthNumber), _
            Left$(MonthName(monthNumber), 3), _
            "'" & MonthName(monthNumber) & " Calendar'!A1"
    Next monthNumber

    With ws.Range("A7:L7")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ws.Rows(7).RowHeight = 24
End Sub

Private Sub StyleCoreWorkbookSheets(ByVal wb As Workbook)
    Dim ws As Worksheet

    Set ws = wb.Worksheets(CALENDAR_DASHBOARD_SHEET)
    ws.Tab.Color = RGB(31, 78, 121)
    ws.Cells.Font.Name = "Aptos"
    ws.Range("B3").NumberFormat = "MM/DD/YYYY HH:MM"
    If ws.Columns("B").ColumnWidth < 20 Then
        ws.Columns("B").ColumnWidth = 20
    End If

    StyleCampaignSheet wb.Worksheets(SH_EMAIL), GetInventoryTable(), _
        RGB(0, 112, 192)
    StyleCampaignSheet wb.Worksheets(SH_SMS), GetSmsTable(), _
        RGB(112, 48, 160)
End Sub

Private Sub StyleCampaignSheet( _
    ByVal ws As Worksheet, _
    ByVal lo As ListObject, _
    ByVal tabColor As Long)

    Dim lc As ListColumn
    Dim stageColumn As ListColumn
    Dim stageCell As Range
    Dim condition As FormatCondition

    ws.Tab.Color = tabColor
    ws.Cells.Font.Name = "Aptos"
    ws.Cells.Font.Strikethrough = False
    ws.Rows(1).RowHeight = 42
    ws.Rows(1).WrapText = True
    ws.Rows(1).VerticalAlignment = xlCenter
    lo.TableStyle = "TableStyleMedium2"

    For Each lc In lo.ListColumns
        Select Case lc.Name
            Case "Send Date"
                lc.Range.ColumnWidth = 12
                lc.DataBodyRange.NumberFormat = "MM/DD/YYYY"
            Case "Send Time"
                lc.Range.ColumnWidth = 11
            Case "Campaign Name"
                lc.Range.ColumnWidth = 34
            Case "Campaign Type"
                lc.Range.ColumnWidth = 14
            Case "Current Stage"
                lc.Range.ColumnWidth = 20
            Case "Owner"
                lc.Range.ColumnWidth = 16
            Case "Campaign Name and UTM Parameter (Source Code)"
                lc.Range.ColumnWidth = 22
            Case "Creative Brief, SL & PH"
                lc.Range.ColumnWidth = 20
            Case "SKUs", "Route", "Send Test"
                lc.Range.ColumnWidth = 12
            Case "In-Design", "Build, QA", "Approval", "Segments", _
                "Send SMS Options"
                lc.Range.ColumnWidth = 13
            Case "Jira Link", "ClickUp Link", "Bluecore Link"
                lc.Range.ColumnWidth = 14
            Case "Est. Audience"
                lc.Range.ColumnWidth = 14
            Case "Delivered"
                lc.Range.ColumnWidth = 12
            Case "Last Updated"
                lc.Range.ColumnWidth = 20
                lc.DataBodyRange.NumberFormat = "MM/DD/YYYY HH:MM"
            Case "Last Updated By"
                lc.Range.ColumnWidth = 18
                lc.DataBodyRange.NumberFormat = "@"
        End Select
    Next lc

    Set stageColumn = FindTableColumn(lo, "Current Stage")
    If Not stageColumn Is Nothing Then
        If Not stageColumn.DataBodyRange Is Nothing Then
            stageColumn.DataBodyRange.FormatConditions.Delete

            Set condition = stageColumn.DataBodyRange.FormatConditions.Add( _
                Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Sent""")
            condition.Interior.Color = RGB(198, 239, 206)
            condition.Font.Color = RGB(0, 97, 0)

            Set condition = stageColumn.DataBodyRange.FormatConditions.Add( _
                Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Scheduled""")
            condition.Interior.Color = RGB(221, 235, 247)
            condition.Font.Color = RGB(31, 78, 121)

            Set condition = stageColumn.DataBodyRange.FormatConditions.Add( _
                Type:=xlCellValue, Operator:=xlEqual, _
                Formula1:="=""Ready to Schedule""")
            condition.Interior.Color = RGB(255, 235, 156)
            condition.Font.Color = RGB(156, 101, 0)

            On Error Resume Next
            For Each stageCell In stageColumn.DataBodyRange.Cells
                stageCell.Errors.item(xlEmptyCellReferences).Ignore = True
            Next stageCell
            On Error GoTo 0
        End If
    End If
End Sub

Private Sub OrderWorkbookSheets(ByVal wb As Workbook)
    Dim dashboard As Worksheet
    Dim emailSheet As Worksheet
    Dim smsSheet As Worksheet
    Dim previousSheet As Worksheet
    Dim ws As Worksheet
    Dim monthNumber As Long

    Set dashboard = wb.Worksheets(CALENDAR_DASHBOARD_SHEET)
    Set emailSheet = wb.Worksheets(SH_EMAIL)
    Set smsSheet = wb.Worksheets(SH_SMS)

    dashboard.Move Before:=wb.Worksheets(1)
    emailSheet.Move After:=dashboard
    smsSheet.Move After:=emailSheet
    Set previousSheet = smsSheet

    For monthNumber = 1 To 12
        Set ws = wb.Worksheets(MonthName(monthNumber) & " Calendar")
        ws.Move After:=previousSheet
        Set previousSheet = ws
    Next monthNumber
End Sub

Private Sub ConfigureWorkbookViews(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim monthNumber As Long

    Set ws = wb.Worksheets(CALENDAR_DASHBOARD_SHEET)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
    ActiveWindow.FreezePanes = False
    ws.Range("A11").Select
    ActiveWindow.FreezePanes = True

    Set ws = wb.Worksheets(CALENDAR_INVENTORY_SHEET)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    Set ws = wb.Worksheets(SH_SMS)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 85
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    For monthNumber = 1 To 12
        Set ws = wb.Worksheets(MonthName(monthNumber) & " Calendar")
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.Zoom = 90
        ActiveWindow.FreezePanes = False
        ws.Range("A6").Select
        ActiveWindow.FreezePanes = True
    Next monthNumber

    wb.Worksheets(CALENDAR_DASHBOARD_SHEET).Activate
    wb.Worksheets(CALENDAR_DASHBOARD_SHEET).Range("A1").Select
End Sub

Private Sub SetInternalLink( _
    ByVal ws As Worksheet, _
    ByVal targetCell As Range, _
    ByVal displayText As String, _
    ByVal subAddress As String)

    targetCell.value = displayText
    ws.Hyperlinks.Add _
        anchor:=targetCell, _
        address:="", _
        subAddress:=subAddress, _
        TextToDisplay:=displayText
End Sub

Private Sub DeleteSheetIfPresent( _
    ByVal wb As Workbook, _
    ByVal sheetName As String)

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then ws.Delete
End Sub
