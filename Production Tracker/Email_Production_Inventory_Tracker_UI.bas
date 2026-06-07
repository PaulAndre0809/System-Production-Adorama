Attribute VB_Name = "modEmailProductionInventoryTracker"
Option Explicit

Public Const SH_INVENTORY As String = "Production Inventory"
Public Const SH_DASHBOARD As String = "Dashboard"
Public Const SH_LOG As String = "Automation Log"

Private Const TBL_INVENTORY As String = "ProductionInventoryTable"
Private Const FIRST_DATA_ROW As Long = 2
Private Const FIRST_DASHBOARD_ROW As Long = 11
Private Const LAST_DASHBOARD_ROW As Long = 110
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

    On Error GoTo SafeExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set lo = GetInventoryTable()
    ApplyCalculatedColumns lo
    RefreshDashboard
    LogAction "RefreshProductionStatus", _
        "Production Inventory and Dashboard refreshed"

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
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
    Dim lo As ListObject
    Dim r As Long
    Dim sourceRow As Long
    Dim checkedMark As String
    Dim campaignRef As String
    Dim stageRef As String
    Dim sendDateRef As String
    Dim sourceCodeRef As String
    Dim briefRef As String
    Dim buildQaRef As String
    Dim approvalRef As String

    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set lo = GetInventoryTable()
    checkedMark = CheckedSymbol()

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
        "=COUNTIFS(" & campaignRef & ",""<>""," & _
        stageRef & ",""<>Sent""," & stageRef & ",""<>Cancelled"")"
    wsD.Range("C5").Formula = _
        "=COUNTIFS(" & sendDateRef & ",TODAY()," & campaignRef & ",""<>"")"
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

    For r = FIRST_DASHBOARD_ROW To LAST_DASHBOARD_ROW
        sourceRow = r - 9

        wsD.Cells(r, "A").Formula = MirrorFormula( _
            lo, sourceRow, "Campaign Name", "Send Date")
        wsD.Cells(r, "B").Formula = MirrorTextFormula( _
            lo, sourceRow, "Campaign Name", "Send Time")
        wsD.Cells(r, "C").Formula = MirrorTextFormula( _
            lo, sourceRow, "Campaign Name", "Campaign Name")
        wsD.Cells(r, "D").Formula = MirrorTextFormula( _
            lo, sourceRow, "Campaign Name", "Campaign Type")
        wsD.Cells(r, "E").Formula = MirrorTextFormula( _
            lo, sourceRow, "Campaign Name", "Current Stage")
        wsD.Cells(r, "F").Formula = MirrorTextFormula( _
            lo, sourceRow, "Campaign Name", "Owner")
        wsD.Cells(r, "G").Formula = MirrorTextFormula( _
            lo, sourceRow, "Campaign Name", "SKUs")
        wsD.Cells(r, "H").Formula = MirrorTextFormula( _
            lo, sourceRow, "Campaign Name", "Build, QA")
        wsD.Cells(r, "I").Formula = MirrorTextFormula( _
            lo, sourceRow, "Campaign Name", "Approval")
        wsD.Cells(r, "J").Formula = LinkFormula(lo, sourceRow, "Jira Link", "Jira")
        wsD.Cells(r, "K").Formula = LinkFormula( _
            lo, sourceRow, "ClickUp Link", "ClickUp")
        wsD.Cells(r, "L").Formula = LinkFormula( _
            lo, sourceRow, "Bluecore Link", "Bluecore")
    Next r

    wsD.Range("A11:A110").NumberFormat = "MM/DD/YYYY"
    wsD.Range("G11:I110").Font.Name = "Segoe UI Symbol"
    wsD.Range("G11:I110").HorizontalAlignment = xlCenter

    wsD.Range("A115").Value = "Today's Sends"
    wsD.Range("B115").Value = "Send Date = today"
    wsD.Range("C115").Value = "Confirm scheduled/send status"
    wsD.Range("D115").Formula = _
        "=COUNTIFS(" & sendDateRef & ",TODAY()," & campaignRef & ",""<>"")"

    wsD.Range("A116").Value = "Source Code Pending"
    wsD.Range("B116").Value = "Source-code checkbox is open"
    wsD.Range("C116").Value = "Finish campaign name and UTM parameters"
    wsD.Range("D116").Formula = PendingCountFormula( _
        campaignRef, sourceCodeRef, checkedMark)

    wsD.Range("A117").Value = "Creative Brief Pending"
    wsD.Range("B117").Value = "Creative checkbox is open"
    wsD.Range("C117").Value = "Complete brief, subject line, and preheader"
    wsD.Range("D117").Formula = PendingCountFormula( _
        campaignRef, briefRef, checkedMark)

    wsD.Range("A118").Value = "Build / QA Pending"
    wsD.Range("B118").Value = "Build/QA checkbox is open"
    wsD.Range("C118").Value = "Complete build and QA"
    wsD.Range("D118").Formula = PendingCountFormula( _
        campaignRef, buildQaRef, checkedMark)

    wsD.Range("A119").Value = "Approval Pending"
    wsD.Range("B119").Value = "Approval checkbox is open"
    wsD.Range("C119").Value = "Obtain final approval"
    wsD.Range("D119").Formula = PendingCountFormula( _
        campaignRef, approvalRef, checkedMark)

    LogAction "RefreshDashboard", "Dashboard formulas refreshed"
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
    Dim nextRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    ws.Cells(nextRow, "A").Value = Now
    ws.Cells(nextRow, "B").Value = Environ$("Username")
    ws.Cells(nextRow, "C").Value = actionName
    ws.Cells(nextRow, "D").Value = details
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
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Checklist"
        .InputMessage = "Choose unchecked or checked."
        .ErrorTitle = "Invalid checklist value"
        .ErrorMessage = "Choose a value from the checklist dropdown."
        .ShowInput = True
        .ShowError = True
    End With

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

    formulaText = "=IF([@[Campaign Name]]="""","""","
    formulaText = formulaText & _
        "IF(AND(ISNUMBER([@Delivered]),[@Delivered]>0),""Sent"","
    formulaText = formulaText & _
        "IF([@[Campaign Name and UTM Parameter (Source Code)]]<>""" & _
        checkedMark & """,""Source Code"","
    formulaText = formulaText & _
        "IF([@[Creative Brief, SL & PH]]<>""" & checkedMark & _
        """,""Creative Brief"","
    formulaText = formulaText & _
        "IF([@SKUs]<>""" & checkedMark & """,""Waiting for SKUs"","
    formulaText = formulaText & _
        "IF([@[In-Design]]<>""" & checkedMark & """,""With Design"","
    formulaText = formulaText & _
        "IF([@[Build, QA]]<>""" & checkedMark & """,""Build / QA"","
    formulaText = formulaText & _
        "IF([@Route]<>""" & checkedMark & """,""Routing"","
    formulaText = formulaText & _
        "IF([@Approval]<>""" & checkedMark & """,""Awaiting Approval"","
    formulaText = formulaText & _
        "IF([@Segments]<>""" & checkedMark & """,""Segments"","
    formulaText = formulaText & _
        "IF(OR(IFERROR(LEN(TRIM([@[Jira Link]]&"""")),0)=0," & _
        "IFERROR(LEN(TRIM([@[ClickUp Link]]&"""")),0)=0," & _
        "IFERROR(LEN(TRIM([@[Bluecore Link]]&"""")),0)=0),""Links Pending"","
    formulaText = formulaText & _
        "IF(IFERROR(LEN(TRIM([@[Est. Audience]]&"""")),0)=0," & _
        """Ready to Schedule"",""Scheduled""" & String$(12, ")")

    If Not stageColumn.DataBodyRange Is Nothing Then
        stageColumn.DataBodyRange.Formula = formulaText
    End If

    If Not updatedColumn.DataBodyRange Is Nothing Then
        updatedColumn.DataBodyRange.Formula = _
            "=IF([@[Campaign Name]]="""","""",NOW())"
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

Private Function GetInventoryTable() As ListObject
    Set GetInventoryTable = _
        ThisWorkbook.Worksheets(SH_INVENTORY).ListObjects(TBL_INVENTORY)
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
    Dim value As String

    value = LCase$(Trim$(headerName))
    value = Replace(value, " ", "")
    value = Replace(value, "/", "")
    value = Replace(value, "\", "")
    value = Replace(value, "-", "")
    value = Replace(value, "_", "")
    value = Replace(value, ".", "")
    value = Replace(value, ",", "")
    value = Replace(value, "&", "")
    value = Replace(value, "(", "")
    value = Replace(value, ")", "")

    HeaderKey = value
End Function

Private Function ValueByHeader( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal headerName As String) As Variant

    ValueByHeader = ws.Cells( _
        rowNumber, InventoryColumnNumber(headerName)).Value
End Function

Private Function IsChecked(ByVal value As Variant) As Boolean
    Dim textValue As String

    If VarType(value) = vbBoolean Then
        IsChecked = CBool(value)
        Exit Function
    End If

    If IsNumeric(value) Then
        IsChecked = (CDbl(value) <> 0)
        Exit Function
    End If

    textValue = LCase$(Trim$(CStr(value)))

    IsChecked = _
        (textValue = LCase$(CheckedSymbol())) Or _
        (textValue = "true") Or _
        (textValue = "yes") Or _
        (textValue = "x") Or _
        (textValue = "done") Or _
        (textValue = "complete") Or _
        (textValue = "completed")
End Function

Private Function HasText(ByVal value As Variant) As Boolean
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function

    HasText = (Len(Trim$(CStr(value))) > 0)
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

    Dim columnLetter As String

    columnLetter = TableColumnLetter(lo, headerName)
    FullColumnReference = "'" & lo.Parent.Name & "'!" & _
        columnLetter & ":" & columnLetter
End Function

Private Function TableCellReference( _
    ByVal lo As ListObject, _
    ByVal sourceRow As Long, _
    ByVal headerName As String) As String

    TableCellReference = "'" & lo.Parent.Name & "'!$" & _
        TableColumnLetter(lo, headerName) & sourceRow
End Function

Private Function TableColumnLetter( _
    ByVal lo As ListObject, _
    ByVal headerName As String) As String

    Dim lc As ListColumn
    Dim absoluteColumn As Long
    Dim addressParts As Variant

    Set lc = FindTableColumn(lo, headerName)
    If lc Is Nothing Then
        Err.Raise vbObjectError + 1050, , _
            "Production Inventory column not found: " & headerName
    End If

    absoluteColumn = lo.Range.Column + lc.Index - 1
    addressParts = Split(lo.Parent.Cells(1, absoluteColumn).Address, "$")
    TableColumnLetter = CStr(addressParts(1))
End Function

Private Function PendingCountFormula( _
    ByVal campaignRef As String, _
    ByVal checklistRef As String, _
    ByVal checkedMark As String) As String

    PendingCountFormula = _
        "=COUNTIFS(" & campaignRef & ",""<>""," & _
        checklistRef & ",""<>" & checkedMark & """)"
End Function

Private Function MirrorFormula( _
    ByVal lo As ListObject, _
    ByVal sourceRow As Long, _
    ByVal campaignHeader As String, _
    ByVal outputHeader As String) As String

    MirrorFormula = "=IF(" & _
        TableCellReference(lo, sourceRow, campaignHeader) & _
        "="""",""""," & _
        TableCellReference(lo, sourceRow, outputHeader) & ")"
End Function

Private Function MirrorTextFormula( _
    ByVal lo As ListObject, _
    ByVal sourceRow As Long, _
    ByVal campaignHeader As String, _
    ByVal outputHeader As String) As String

    MirrorTextFormula = "=IF(" & _
        TableCellReference(lo, sourceRow, campaignHeader) & _
        "="""",""""," & _
        TableCellReference(lo, sourceRow, outputHeader) & "&"""")"
End Function

Private Function LinkFormula( _
    ByVal lo As ListObject, _
    ByVal sourceRow As Long, _
    ByVal linkHeader As String, _
    ByVal displayText As String) As String

    Dim linkRef As String

    linkRef = TableCellReference(lo, sourceRow, linkHeader)
    LinkFormula = "=IF(" & linkRef & "="""","""",HYPERLINK(" & _
        linkRef & ",""" & displayText & """))"
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

'Optional workbook event. Place this in ThisWorkbook:
'
'Private Sub Workbook_Open()
'    RefreshProductionStatus
'End Sub
'
'Optional sheet event. Place this in the Production Inventory sheet module:
'
'Private Sub Worksheet_Change(ByVal Target As Range)
'    On Error GoTo SafeExit
'    If Intersect(Target, Me.ListObjects("ProductionInventoryTable").Range) _
'        Is Nothing Then Exit Sub
'    Application.EnableEvents = False
'    RefreshProductionStatus
'SafeExit:
'    Application.EnableEvents = True
'End Sub

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
    OrderWorkbookSheets wb
    ConfigureWorkbookViews wb

    wb.Worksheets("Dropdowns").Visible = xlSheetVeryHidden
    wb.Worksheets("Automation Log").Visible = xlSheetVeryHidden

    Application.CalculateFull
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
    ws.Range("A1").Formula2 = _
        "=TEXT(DATE(YEAR(TODAY())," & monthNumber & _
        ",1),""mmmm yyyy"")&"" Campaign Calendar"""

    ws.Range("A2:G2").Merge
    ws.Range("A2").Value = _
        "Campaigns update automatically from Production Inventory send dates."

    SetInternalLink ws, ws.Range("A3"), _
        "Dashboard", "'Dashboard'!A1"
    SetInternalLink ws, ws.Range("B3"), _
        "Production Inventory", "'Production Inventory'!A1"

    ws.Range("C3:E3").Merge
    ws.Range("C3").Formula2 = _
        "=""Today: ""&TEXT(TODAY(),""MM/DD/YYYY"")"

    SetInternalLink ws, ws.Range("F3"), _
        "< " & MonthName(previousMonth), _
        "'" & MonthName(previousMonth) & " Calendar'!A1"
    SetInternalLink ws, ws.Range("G3"), _
        MonthName(nextMonth) & " >", _
        "'" & MonthName(nextMonth) & " Calendar'!A1"

    ws.Range("A4:G4").Merge
    ws.Range("A4").Value = _
        "Each campaign appears on its scheduled Send Date. " & _
        "Multiple campaigns are listed together."

    dayNames = Array( _
        "Sunday", "Monday", "Tuesday", "Wednesday", _
        "Thursday", "Friday", "Saturday")

    For dayIndex = 0 To 6
        ws.Cells(5, dayIndex + 1).Value = dayNames(dayIndex)
    Next dayIndex

    formulaText = _
        "=LET(first,DATE(YEAR(TODAY())," & monthNumber & ",1)," & _
        "d,first-WEEKDAY(first,1)+1+(ROW()-6)*7+COLUMN()-1," & _
        "items,TEXTJOIN(CHAR(10)&""" & bullet & " "",TRUE," & _
        "FILTER(" & CALENDAR_TABLE & "[Campaign Name]," & _
        "IFERROR(INT(" & CALENDAR_TABLE & "[Send Date]),0)=d,""""))," & _
        "IF(MONTH(d)<>" & monthNumber & ",""""," & _
        "DAY(d)&IF(items="""","""",CHAR(10)&""" & bullet & " ""&items)))"

    Set calendarRange = ws.Range( _
        "A" & FIRST_CALENDAR_ROW & ":G" & LAST_CALENDAR_ROW)
    calendarRange.Formula2 = formulaText

    ws.Range("A12:G12").Merge
    ws.Range("A12").Value = _
        "Update Send Date or Campaign Name in Production Inventory " & _
        "to refresh this calendar."

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

    ws.Range("A3:B3").Interior.Color = RGB(221, 235, 247)
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

    If monthNumber = Month(Date) Then
        ws.Tab.Color = RGB(255, 235, 156)
    Else
        ws.Tab.Color = RGB(91, 155, 213)
    End If

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
    Dim lo As ListObject
    Dim lc As ListColumn

    Set ws = wb.Worksheets(CALENDAR_DASHBOARD_SHEET)
    ws.Tab.Color = RGB(31, 78, 121)
    ws.Range("B3").NumberFormat = "MM/DD/YYYY HH:MM"
    If ws.Columns("B").ColumnWidth < 20 Then
        ws.Columns("B").ColumnWidth = 20
    End If

    Set ws = wb.Worksheets(CALENDAR_INVENTORY_SHEET)
    ws.Tab.Color = RGB(0, 112, 192)
    ws.Rows(1).RowHeight = 42
    ws.Rows(1).WrapText = True
    ws.Rows(1).VerticalAlignment = xlCenter

    Set lo = ws.ListObjects(CALENDAR_TABLE)

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
            Case "SKUs", "Route"
                lc.Range.ColumnWidth = 12
            Case "In-Design", "Build, QA", "Approval", "Segments"
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
        End Select
    Next lc
End Sub

Private Sub OrderWorkbookSheets(ByVal wb As Workbook)
    Dim dashboard As Worksheet
    Dim inventory As Worksheet
    Dim previousSheet As Worksheet
    Dim ws As Worksheet
    Dim monthNumber As Long

    Set dashboard = wb.Worksheets(CALENDAR_DASHBOARD_SHEET)
    Set inventory = wb.Worksheets(CALENDAR_INVENTORY_SHEET)

    dashboard.Move Before:=wb.Worksheets(1)
    inventory.Move After:=dashboard
    Set previousSheet = inventory

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

    targetCell.Value = displayText
    ws.Hyperlinks.Add _
        Anchor:=targetCell, _
        Address:="", _
        SubAddress:=subAddress, _
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
