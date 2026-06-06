Attribute VB_Name = "modEmailProductionTracker"
Option Explicit

Public Const SH_INVENTORY As String = "Production Inventory"
Public Const SH_DASHBOARD As String = "Dashboard"
Public Const SH_LOG As String = "Automation Log"

'Run after updating campaign rows.
Public Sub RefreshProductionStatus()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = ThisWorkbook.Worksheets(SH_INVENTORY)
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = 2 To lastRow
        If Trim(ws.Cells(r, "D").Value) <> "" Then
            ws.Cells(r, "A").Value = GenerateCampaignID(ws.Cells(r, "B").Value, CStr(ws.Cells(r, "D").Value))
            ws.Cells(r, "F").Value = CalculateCurrentStage(ws, r)
            ws.Cells(r, "V").Value = CalculateRiskLevel(ws, r)
            ws.Cells(r, "W").Value = Now
        Else
            ws.Cells(r, "A").ClearContents
            ws.Cells(r, "F").ClearContents
            ws.Cells(r, "V").ClearContents
            ws.Cells(r, "W").ClearContents
        End If
    Next r

    RefreshDashboard
    LogAction "RefreshProductionStatus", "Production Inventory and Dashboard refreshed"

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Function GenerateCampaignID(ByVal sendDate As Variant, ByVal campaignName As String) As String
    Dim cleanName As String

    cleanName = UCase(campaignName)
    cleanName = Replace(cleanName, "Sample | ", "")
    cleanName = Replace(cleanName, " ", "-")
    cleanName = Replace(cleanName, "/", "-")
    cleanName = Replace(cleanName, "&", "AND")
    cleanName = Replace(cleanName, "--", "-")

    If IsDate(sendDate) Then
        GenerateCampaignID = "EMAIL-" & Format(CDate(sendDate), "yyyy-mm-dd") & "-" & Left(cleanName, 25)
    Else
        GenerateCampaignID = "EMAIL-TBD-" & Left(cleanName, 25)
    End If
End Function

Public Function CalculateCurrentStage(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim skuStatus As String
    Dim briefStatus As String
    Dim designStatus As String
    Dim builderStatus As String
    Dim segmentStatus As String
    Dim qaStatus As String
    Dim scheduleStatus As String
    Dim delivered As Variant

    skuStatus = Trim(ws.Cells(r, "J").Value)
    briefStatus = Trim(ws.Cells(r, "K").Value)
    designStatus = Trim(ws.Cells(r, "L").Value)
    builderStatus = Trim(ws.Cells(r, "M").Value)
    segmentStatus = Trim(ws.Cells(r, "N").Value)
    qaStatus = Trim(ws.Cells(r, "O").Value)
    scheduleStatus = Trim(ws.Cells(r, "P").Value)
    delivered = ws.Cells(r, "U").Value

    If scheduleStatus = "Cancelled" Then
        CalculateCurrentStage = "Cancelled"
    ElseIf IsNumeric(delivered) And delivered > 0 Then
        CalculateCurrentStage = "Sent"
    ElseIf scheduleStatus = "Scheduled" Then
        CalculateCurrentStage = "Scheduled"
    ElseIf qaStatus = "Passed" Then
        CalculateCurrentStage = "QA Passed"
    ElseIf qaStatus = "In Progress" Then
        CalculateCurrentStage = "QA in Progress"
    ElseIf segmentStatus = "Ready" Or segmentStatus = "Created" Then
        CalculateCurrentStage = "Segment Ready"
    ElseIf ws.Cells(r, "S").Value <> "" Then
        CalculateCurrentStage = "Built in Bluecore"
    ElseIf builderStatus = "Sent" Then
        CalculateCurrentStage = "Sent to Builder"
    ElseIf designStatus = "Ready" Then
        CalculateCurrentStage = "Design Ready"
    ElseIf designStatus = "In Progress" Then
        CalculateCurrentStage = "With Design"
    ElseIf briefStatus = "Ready" Or briefStatus = "Approved" Then
        CalculateCurrentStage = "Brief Ready"
    ElseIf briefStatus = "In Progress" Or briefStatus = "Draft" Then
        CalculateCurrentStage = "Brief in Progress"
    ElseIf skuStatus = "Needed" Or skuStatus = "Requested" Then
        CalculateCurrentStage = "Waiting for SKUs"
    ElseIf Trim(ws.Cells(r, "H").Value) <> "" Then
        CalculateCurrentStage = "Blocked"
    Else
        CalculateCurrentStage = "Not Started"
    End If
End Function

Public Function CalculateRiskLevel(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim sendDate As Variant
    Dim currentStage As String
    Dim blocker As String
    Dim daysLeft As Long

    sendDate = ws.Cells(r, "B").Value
    currentStage = ws.Cells(r, "F").Value
    blocker = Trim(ws.Cells(r, "H").Value)

    If currentStage = "Sent" Or currentStage = "Cancelled" Then
        CalculateRiskLevel = "Complete"
        Exit Function
    End If

    If currentStage = "Scheduled" Then
        CalculateRiskLevel = "Ready"
        Exit Function
    End If

    If currentStage = "Blocked" Or blocker <> "" Then
        CalculateRiskLevel = "Blocked"
        Exit Function
    End If

    If Not IsDate(sendDate) Then
        CalculateRiskLevel = "At Risk"
        Exit Function
    End If

    daysLeft = DateDiff("d", Date, CDate(sendDate))

    If daysLeft < 0 Then
        CalculateRiskLevel = "Critical"
    ElseIf daysLeft = 0 And currentStage <> "Scheduled" Then
        CalculateRiskLevel = "Critical"
    ElseIf daysLeft = 1 And currentStage <> "QA Passed" And currentStage <> "Scheduled" Then
        CalculateRiskLevel = "High Risk"
    ElseIf daysLeft <= 2 And currentStage <> "Scheduled" Then
        CalculateRiskLevel = "At Risk"
    Else
        CalculateRiskLevel = "On Track"
    End If
End Function

'Refreshes the new Dashboard UI.
Public Sub RefreshDashboard()
    Dim wsD As Worksheet
    Dim r As Long
    Dim sourceRow As Long

    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)

    'KPI cards
    wsD.Range("B3").Formula = "=NOW()"
    wsD.Range("A5").Formula = "=COUNTIFS('" & SH_INVENTORY & "'!D:D,""<>"",'" & SH_INVENTORY & "'!F:F,""<>Sent"",'" & SH_INVENTORY & "'!F:F,""<>Cancelled"")"
    wsD.Range("C5").Formula = "=COUNTIFS('" & SH_INVENTORY & "'!B:B,TODAY(),'" & SH_INVENTORY & "'!D:D,""<>"")"
    wsD.Range("E5").Formula = "=COUNTIFS('" & SH_INVENTORY & "'!V:V,""Blocked"")"
    wsD.Range("G5").Formula = "=COUNTIFS('" & SH_INVENTORY & "'!V:V,""At Risk"")+COUNTIFS('" & SH_INVENTORY & "'!V:V,""High Risk"")+COUNTIFS('" & SH_INVENTORY & "'!V:V,""Critical"")"
    wsD.Range("I5").Formula = "=COUNTIFS('" & SH_INVENTORY & "'!F:F,""Scheduled"")"
    wsD.Range("K5").Formula = "=COUNTIFS('" & SH_INVENTORY & "'!F:F,""Sent"")"

    'All Work in Production mirror table
    For r = 11 To 110
        sourceRow = r - 9
        wsD.Cells(r, "A").Formula = "=IF('" & SH_INVENTORY & "'!$D" & sourceRow & "="""","""",'" & SH_INVENTORY & "'!$B" & sourceRow & ")"
        wsD.Cells(r, "B").Formula = "=IF('" & SH_INVENTORY & "'!$D" & sourceRow & "="""","""",'" & SH_INVENTORY & "'!$C" & sourceRow & "&"")"
        wsD.Cells(r, "C").Formula = "=IF('" & SH_INVENTORY & "'!$D" & sourceRow & "="""","""",'" & SH_INVENTORY & "'!$D" & sourceRow & "&"")"
        wsD.Cells(r, "D").Formula = "=IF('" & SH_INVENTORY & "'!$D" & sourceRow & "="""","""",'" & SH_INVENTORY & "'!$E" & sourceRow & "&"")"
        wsD.Cells(r, "E").Formula = "=IF('" & SH_INVENTORY & "'!$D" & sourceRow & "="""","""",'" & SH_INVENTORY & "'!$F" & sourceRow & "&"")"
        wsD.Cells(r, "F").Formula = "=IF('" & SH_INVENTORY & "'!$D" & sourceRow & "="""","""",'" & SH_INVENTORY & "'!$G" & sourceRow & "&"")"
        wsD.Cells(r, "G").Formula = "=IF('" & SH_INVENTORY & "'!$D" & sourceRow & "="""","""",'" & SH_INVENTORY & "'!$V" & sourceRow & "&"")"
        wsD.Cells(r, "H").Formula = "=IF('" & SH_INVENTORY & "'!$D" & sourceRow & "="""","""",'" & SH_INVENTORY & "'!$H" & sourceRow & "&"")"
        wsD.Cells(r, "I").Formula = "=IF('" & SH_INVENTORY & "'!$D" & sourceRow & "="""","""",'" & SH_INVENTORY & "'!$I" & sourceRow & "&"")"
        wsD.Cells(r, "J").Formula = "=IF('" & SH_INVENTORY & "'!$Q" & sourceRow & "="""","""",HYPERLINK('" & SH_INVENTORY & "'!$Q" & sourceRow & ",""Jira""))"
        wsD.Cells(r, "K").Formula = "=IF('" & SH_INVENTORY & "'!$R" & sourceRow & "="""","""",HYPERLINK('" & SH_INVENTORY & "'!$R" & sourceRow & ",""ClickUp""))"
        wsD.Cells(r, "L").Formula = "=IF('" & SH_INVENTORY & "'!$S" & sourceRow & "="""","""",HYPERLINK('" & SH_INVENTORY & "'!$S" & sourceRow & ",""Bluecore""))"
    Next r

    'Daily Action Lanes counts
    wsD.Range("D115").Formula = "=COUNTIFS('" & SH_INVENTORY & "'!B:B,TODAY(),'" & SH_INVENTORY & "'!D:D,""<>"")"
    wsD.Range("D116").Formula = "=COUNTIFS('" & SH_INVENTORY & "'!V:V,""Blocked"")"
    wsD.Range("D117").Formula = "=COUNTIFS('" & SH_INVENTORY & "'!L:L,""Ready"",'" & SH_INVENTORY & "'!R:R,"""")"
    wsD.Range("D118").Formula = "=COUNTIFS('" & SH_INVENTORY & "'!M:M,""Sent"",'" & SH_INVENTORY & "'!S:S,"""")"
    wsD.Range("D119").Formula = "=COUNTIFS('" & SH_INVENTORY & "'!O:O,""Passed"",'" & SH_INVENTORY & "'!P:P,""<>Scheduled"")"

    LogAction "RefreshDashboard", "Dashboard UI formulas refreshed"
End Sub

Public Sub CreateDailyDigest()
    Dim ws As Worksheet
    Dim wsD As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim digest As String

    Set ws = ThisWorkbook.Worksheets(SH_INVENTORY)
    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

    digest = "EMAIL PRODUCTION DIGEST - " & Format(Date, "mmmm d, yyyy") & vbCrLf & vbCrLf

    digest = digest & "TODAY'S SENDS:" & vbCrLf
    For r = 2 To lastRow
        If IsDate(ws.Cells(r, "B").Value) Then
            If CDate(ws.Cells(r, "B").Value) = Date Then
                digest = digest & "- " & ws.Cells(r, "D").Value & " | " & ws.Cells(r, "C").Text & " | " & ws.Cells(r, "F").Value & vbCrLf
            End If
        End If
    Next r

    digest = digest & vbCrLf & "AT RISK / HIGH RISK / CRITICAL:" & vbCrLf
    For r = 2 To lastRow
        If ws.Cells(r, "V").Value = "At Risk" Or ws.Cells(r, "V").Value = "High Risk" Or ws.Cells(r, "V").Value = "Critical" Then
            digest = digest & "- " & ws.Cells(r, "D").Value & " | " & ws.Cells(r, "V").Value & " | Next: " & ws.Cells(r, "I").Value & vbCrLf
        End If
    Next r

    digest = digest & vbCrLf & "BLOCKED:" & vbCrLf
    For r = 2 To lastRow
        If ws.Cells(r, "V").Value = "Blocked" Then
            digest = digest & "- " & ws.Cells(r, "D").Value & " | Blocker: " & ws.Cells(r, "H").Value & vbCrLf
        End If
    Next r

    wsD.Range("N3").Value = "Daily Digest"
    wsD.Range("N4").Value = digest
    wsD.Range("N4").WrapText = True
    wsD.Columns("N:N").ColumnWidth = 90

    LogAction "CreateDailyDigest", "Daily digest created"
    MsgBox "Daily digest created on Dashboard sheet, starting at N3.", vbInformation
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

'Optional workbook event:
'Paste this into ThisWorkbook, not into this module:
'
'Private Sub Workbook_Open()
'    RefreshProductionStatus
'End Sub
'
'Optional sheet event:
'Paste this into the Production Inventory sheet code area:
'
'Private Sub Worksheet_Change(ByVal Target As Range)
'    On Error GoTo SafeExit
'    If Intersect(Target, Me.Range("B:V")) Is Nothing Then Exit Sub
'    Application.EnableEvents = False
'    RefreshProductionStatus
'SafeExit:
'    Application.EnableEvents = True
'End Sub
