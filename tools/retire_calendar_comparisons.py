"""Retire calendar sheets and weekly comparison blocks from tracker workbooks.

This is a development utility. The finished XLSM files do not depend on it.
"""

from __future__ import annotations

import argparse
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client as win32


VBA_MODULE = "modEmailProductionTracker"
NOTES_SHEET = "Notes - Instructions"
DASHBOARD_SHEET = "Dashboard"


def normalize_vba(text: str) -> str:
    """Normalize VBA line endings so procedure matching remains predictable."""
    return text.replace("\r\n", "\n").replace("\r", "\n")


def sanitize_vba_continuations(text: str) -> str:
    """Remove blank lines that invalidate VBA continuation statements."""
    lines = normalize_vba(text).split("\n")
    sanitized: list[str] = []
    index = 0

    while index < len(lines):
        line = lines[index]
        sanitized.append(line)
        if line.rstrip().endswith("_"):
            while index + 1 < len(lines) and not lines[index + 1].strip():
                index += 1
        index += 1

    return "\n".join(sanitized)


def procedure_pattern(name: str) -> re.Pattern[str]:
    """Build a case-insensitive pattern for one complete VBA procedure."""
    return re.compile(
        r"(?ims)^[ \t]*(?:Public|Private|Friend)?[ \t]+"
        r"(?:Sub|Function)[ \t]+"
        + re.escape(name)
        + r"\b.*?^[ \t]*End (?:Sub|Function)[ \t]*$"
    )


def replace_procedure(text: str, name: str, replacement: str) -> str:
    """Replace exactly one VBA procedure and fail loudly if it is missing."""
    pattern = procedure_pattern(name)
    updated, count = pattern.subn(replacement.strip(), text, count=1)
    if count != 1:
        raise RuntimeError(f"Expected one VBA procedure named {name}; found {count}.")
    return updated


def transform_procedure(text: str, name: str, transform) -> str:
    """Apply a focused transformation within one VBA procedure."""
    pattern = procedure_pattern(name)
    match = pattern.search(text)
    if match is None:
        raise RuntimeError(f"VBA procedure not found: {name}")
    replacement = transform(match.group(0))
    return text[: match.start()] + replacement + text[match.end() :]


VALIDATE_WORKBOOK = r'''
Public Function ValidateWorkbookConfiguration() As String
    Dim emailTable As ListObject
    Dim smsTable As ListObject
    Dim dashboardTable As ListObject
    Dim dashboardSheet As Worksheet
    Dim emailTable As ListObject
    Dim smsTable As ListObject
    Dim candidateSheet As Worksheet
    Dim candidateTable As ListObject
    Dim candidateChart As ChartObject
    Dim foundColumn As ListColumn
    Dim item As Variant
    Dim emailHeaders As Variant
    Dim smsHeaders As Variant

    On Error GoTo ValidationFailed

    ' Validate the two source tables and the active Dashboard feed.
    Set emailTable = GetInventoryTable()
    Set smsTable = GetSmsTable()
    Set dashboardSheet = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set dashboardTable = dashboardSheet.ListObjects(TBL_DASHBOARD)

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
        "Bluecore/Attentive Link", _
        "Est. Audience", _
        "Delivered", _
        "Last Updated", _
        "Last Updated By", _
        "Notes")

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
        "Proof of Schedule", _
        "Bluecore/Attentive Link", _
        "Est. Audience", _
        "Delivered", _
        "Last Updated", _
        "Last Updated By", _
        "Notes")

    ' Keep embedded validation fast; extensive behavioral QA runs externally.
    For Each item In emailHeaders
        Set foundColumn = Nothing
        Set foundColumn = FindTableColumn(emailTable, CStr(item))
        If foundColumn Is Nothing Then
            Err.Raise vbObjectError + 1080, , _
                "Missing Email Campaigns column: " & CStr(item)
        End If
    Next item

    For Each item In smsHeaders
        Set foundColumn = Nothing
        Set foundColumn = FindTableColumn(smsTable, CStr(item))
        If foundColumn Is Nothing Then
            Err.Raise vbObjectError + 1080, , _
                "Missing SMS Campaigns column: " & CStr(item)
        End If
    Next item

    Set foundColumn = FindTableColumn(emailTable, "Current Stage")
    If Not foundColumn.DataBodyRange.Cells(1, 1).HasFormula Then
        Err.Raise vbObjectError + 1087, , _
            "Email Campaigns Current Stage is not formula-driven."
    End If

    Set foundColumn = FindTableColumn(smsTable, "Current Stage")
    If Not foundColumn.DataBodyRange.Cells(1, 1).HasFormula Then
        Err.Raise vbObjectError + 1087, , _
            "SMS Campaigns Current Stage is not formula-driven."
    End If

    If dashboardTable.ListColumns.Count <> 12 Or _
        dashboardTable.ListRows.Count < 1 Then
        Err.Raise vbObjectError + 1081, , _
            "Dashboard work table has an invalid structure."
    End If

    ' Calendar worksheets are retired and must not be recreated.
    For Each candidateSheet In ThisWorkbook.Worksheets
        If IsLegacyCalendarSheetName(candidateSheet.Name) Then
            Err.Raise vbObjectError + 1087, , _
                "Retired calendar sheet still exists: " & candidateSheet.Name
        End If
    Next candidateSheet

    ' Weekly delivered comparison tables and charts are also retired.
    For Each candidateTable In dashboardSheet.ListObjects
        If InStr(1, candidateTable.Name, "DeliveredComparison", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 1089, , _
                "Retired Dashboard comparison table still exists."
        End If
    Next candidateTable

    For Each candidateChart In dashboardSheet.ChartObjects
        If InStr(1, candidateChart.Name, "Delivered", vbTextCompare) > 0 And _
            InStr(1, candidateChart.Name, "Comparison", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 1091, , _
                "Retired Dashboard comparison chart still exists."
        End If
    Next candidateChart

    If Application.WorksheetFunction.CountA(dashboardSheet.Range("N2:S44")) > 0 Then
        Err.Raise vbObjectError + 1092, , _
            "Retired Dashboard comparison cells were not cleared."
    End If

    If Application.WorksheetFunction.CountA(dashboardSheet.Range("A7:L7")) > 0 Or _
        dashboardSheet.Range("A7:L7").Hyperlinks.Count > 0 Then
        Err.Raise vbObjectError + 1094, , _
            "Retired calendar navigation links were not cleared."
    End If

    If Not InstructionSheetIsReady() Then
        Err.Raise vbObjectError + 1093, , _
            "Notes - Instructions sheet is missing or invalid."
    End If

    ValidateWorkbookConfiguration = "OK"
    Exit Function

ValidationFailed:
    ValidateWorkbookConfiguration = "QA failed: " & Err.Description
End Function
'''


FAST_VALIDATE_WORKBOOK = r'''
Public Function ValidateWorkbookConfiguration() As String
    Dim dashboardSheet As Worksheet
    Dim emailTable As ListObject
    Dim smsTable As ListObject
    Dim candidateSheet As Worksheet
    Dim candidateTable As ListObject
    Dim candidateChart As ChartObject

    On Error GoTo ValidationFailed

    ' Confirm that the active core sheets and tables still exist.
    Set dashboardSheet = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set emailTable = ThisWorkbook.Worksheets(SH_EMAIL).ListObjects(TBL_INVENTORY)
    Set smsTable = ThisWorkbook.Worksheets(SH_SMS).ListObjects(TBL_SMS)

    If dashboardSheet.ListObjects(TBL_DASHBOARD).ListColumns.Count <> 12 Then
        Err.Raise vbObjectError + 1081, , _
            "Dashboard work table has an invalid structure."
    End If

    ' Confirm that retired Calendar sheets cannot reappear unnoticed.
    For Each candidateSheet In ThisWorkbook.Worksheets
        If InStr(1, candidateSheet.Name, "Calendar", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 1087, , _
                "Retired calendar sheet still exists: " & candidateSheet.Name
        End If
    Next candidateSheet

    ' Confirm that retired comparison objects and display cells remain absent.
    For Each candidateTable In dashboardSheet.ListObjects
        If InStr(1, candidateTable.Name, "DeliveredComparison", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 1089, , _
                "Retired Dashboard comparison table still exists."
        End If
    Next candidateTable

    For Each candidateChart In dashboardSheet.ChartObjects
        If InStr(1, candidateChart.Name, "Delivered", vbTextCompare) > 0 And _
            InStr(1, candidateChart.Name, "Comparison", vbTextCompare) > 0 Then
            Err.Raise vbObjectError + 1091, , _
                "Retired Dashboard comparison chart still exists."
        End If
    Next candidateChart

    If Application.WorksheetFunction.CountA(dashboardSheet.Range("N2:S44")) > 0 Or _
        Application.WorksheetFunction.CountA(dashboardSheet.Range("A7:L7")) > 0 Then
        Err.Raise vbObjectError + 1092, , _
            "Retired Dashboard display cells were not cleared."
    End If

    If CStr(ThisWorkbook.Worksheets(SH_INSTRUCTIONS).Range("A1").Value) <> _
        "Detailed Notes and Instructions" Then
        Err.Raise vbObjectError + 1093, , _
            "Notes - Instructions sheet is missing or invalid."
    End If

    ValidateWorkbookConfiguration = "OK"
    Exit Function

ValidationFailed:
    ValidateWorkbookConfiguration = "QA failed: " & Err.Description
End Function
'''


APPLY_ALL_CONFIGURATIONS = r'''
Public Sub ApplyAllConfigurations()
    ' Compatibility no-op: use RefreshDashboard for the active refresh path.
End Sub
'''


REFRESH_DASHBOARD_LIGHT = r'''
Public Sub RefreshDashboard()
    Dim wsD As Worksheet
    Dim dashboardTable As ListObject

    On Error GoTo RefreshExit

    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set dashboardTable = wsD.ListObjects(TBL_DASHBOARD)

    ' Retain the native SharePoint-compatible formula feed and recalculate it.
    wsD.Range("A1").Value = "Email & SMS Campaign Command Center"
    wsD.Range("A2").Value = _
        "Current week through next week Email and SMS campaign command center."
    wsD.Range("A8").Value = "Current Week + Next Week Campaigns"
    wsD.Range("A9").Value = _
        "All campaigns scheduled from this Sunday through next Saturday."

    wsD.Range("AA11").Calculate
    If Not dashboardTable.DataBodyRange Is Nothing Then
        dashboardTable.DataBodyRange.Calculate
    End If

    wsD.Range("A5:K5").Calculate
    wsD.Range("B3:D3").Calculate

RefreshExit:
End Sub
'''


VALIDATE_CAMPAIGN_TABLE = r'''
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
    Dim sampleCell As Range
    Dim controlType As Long

    ' Header validation is inexpensive and protects every downstream formula.
    For Each item In requiredHeaders
        Set lc = FindTableColumn(lo, CStr(item))
        If lc Is Nothing Then
            Err.Raise vbObjectError + 1080, , _
                "Missing " & lo.Parent.Name & " column: " & CStr(item)
        End If
    Next item

    ' Inspect representative cells instead of entire thousand-row columns.
    Set ownerCol = FindTableColumn(lo, "Owner")
    Set sampleCell = ownerCol.DataBodyRange.Cells(1, 1)
    If RangeHasValidation(sampleCell) Then
        Err.Raise vbObjectError + 1083, , _
            lo.Parent.Name & " Owner must be plain text."
    End If

    Set lc = FindTableColumn(lo, "Campaign Type")
    Set sampleCell = lc.DataBodyRange.Cells(1, 1)
    If lc Is Nothing Or Not CampaignTypeDropdownIsConfigured(sampleCell) Then
        Err.Raise vbObjectError + 1092, , _
            lo.Parent.Name & " Campaign Type dropdown is missing or invalid."
    End If

    For Each item In checklistHeaders
        Set lc = FindTableColumn(lo, CStr(item))
        Set sampleCell = lc.DataBodyRange.Cells(1, 1)

        If RangeHasValidation(sampleCell) Then
            Err.Raise vbObjectError + 1084, , _
                "Checklist column contains a dropdown: " & CStr(item)
        End If

        controlType = CheckboxControlType(sampleCell)
        If controlType <> 2 And _
            InStr(1, sampleCell.NumberFormat, _
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
        Not stageCol.DataBodyRange.Cells(1, 1).HasFormula Then
        Err.Raise vbObjectError + 1087, , _
            lo.Parent.Name & " Current Stage is not formula-driven."
    End If
End Sub
'''


COUNT_BROKEN_REFERENCES = r'''
Private Function CountBrokenReferences(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim brokenCell As Range
    Dim nm As Name
    Dim total As Long

    ' Use Excel's native Find engine instead of inspecting every formula cell.
    For Each ws In wb.Worksheets
        Set brokenCell = Nothing
        On Error Resume Next
        Set brokenCell = ws.UsedRange.Find( _
            What:="#REF!", _
            After:=ws.UsedRange.Cells(1, 1), _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False)
        On Error GoTo 0

        If Not brokenCell Is Nothing Then total = total + 1
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
'''


GET_INVENTORY_TABLE = r'''
Private Function GetInventoryTable() As ListObject
    ' Return the Email Campaigns table using standard collection syntax.
    Set GetInventoryTable = _
        ThisWorkbook.Worksheets(SH_EMAIL).ListObjects(TBL_INVENTORY)
End Function
'''


GET_SMS_TABLE = r'''
Private Function GetSmsTable() As ListObject
    ' Return the SMS Campaigns table using standard collection syntax.
    Set GetSmsTable = _
        ThisWorkbook.Worksheets(SH_SMS).ListObjects(TBL_SMS)
End Function
'''


HEADER_KEY = r'''
Private Function HeaderKey(ByVal headerName As String) As String
    ' Exact workbook headers only need case-insensitive, trimmed matching.
    NormalizeHeaderKey = LCase$(Trim$(headerName))
End Function
'''


IS_USER_EDITABLE_COLUMN = r'''
Private Function IsUserEditableColumn( _
    ByVal lo As ListObject, _
    ByVal absoluteColumn As Long) As Boolean

    Dim lc As ListColumn
    Dim headerName As String

    For Each lc In lo.ListColumns
        If lo.Range.Column + lc.Index - 1 = absoluteColumn Then
            headerName = Trim$(lc.Name)
            Exit For
        End If
    Next lc

    Select Case LCase$(headerName)
        Case "current stage", "last updated", "last updated by"
            IsUserEditableColumn = False
        Case Else
            IsUserEditableColumn = (Len(headerName) > 0)
    End Select
End Function
'''


FIND_TABLE_COLUMN = r'''
Private Function FindTableColumn( _
    ByVal lo As ListObject, _
    ByVal headerName As String) As ListColumn

    Dim lc As ListColumn

    ' Match the workbook's exact headers without a separate normalizer.
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(headerName), vbTextCompare) = 0 Then
            Set FindTableColumn = lc
            Exit Function
        End If
    Next lc
End Function
'''


REFRESH_NATIVE_OUTPUTS = r'''
Public Sub RefreshNativeOutputs()
    Dim wsD As Worksheet
    On Error GoTo NativeRefreshExit

    ' Keep edit-time refreshes light; full Dashboard output has its own refresh.
    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)

    wsD.Range("B3:D3").Calculate
    wsD.Range("A5:K5").Calculate

NativeRefreshExit:
End Sub
'''


FORMAT_CHANGED_HYPERLINKS = r'''
Private Sub FormatHyperlinksInChangedCells( _
    ByVal lo As ListObject, _
    ByVal changedCells As Range)

    Dim cell As Range
    Dim colName As String
    Dim displayValue As String
    Dim rawValue As String

    On Error Resume Next

    ' Convert supported pasted URLs without writing debug rows for every edit.
    For Each cell In changedCells.Cells
        If Not IsEmpty(cell.Value) Then
            If Not cell.HasFormula Or _
                InStr(1, cell.Formula, "HYPERLINK", vbTextCompare) = 0 Then

                rawValue = Trim$(CStr(cell.Value))
                If LCase$(Left$(rawValue, 4)) = "http" Then
                    colName = Trim$(CStr(lo.HeaderRowRange.Cells( _
                        1, cell.Column - lo.Range.Column + 1).Value))
                    displayValue = vbNullString

                    Select Case colName
                        Case "Jira Link"
                            displayValue = "Jira"
                        Case "ClickUp Link"
                            displayValue = "ClickUp"
                        Case "Bluecore/Attentive Link", "Bluecore Link"
                            displayValue = "Bluecore/Attentive"
                        Case "Proof of Schedule"
                            displayValue = "Proof of Schedule"
                    End Select

                    If Len(displayValue) > 0 Then
                        cell.Formula = "=HYPERLINK(""" & rawValue & _
                            """,""" & displayValue & """)"
                    End If
                End If
            End If
        End If
    Next cell

    On Error GoTo 0
End Sub
'''


HANDLE_CAMPAIGN_CHANGE = r'''
Public Sub HandleCampaignChange( _
    ByVal ws As Worksheet, _
    ByVal Target As Range)

    Dim lo As ListObject
    Dim changedCells As Range
    Dim rowsToStamp As Object
    Dim cell As Range
    Dim rowKey As Variant
    Dim tableRowIndex As Long
    Dim tableColumnIndex As Long
    Dim headerName As String
    Dim currentUser As String

    On Error GoTo ChangeExit

    Select Case ws.Name
        Case SH_EMAIL
            Set lo = ThisWorkbook.Worksheets(SH_EMAIL) _
                .ListObjects(TBL_INVENTORY)
        Case SH_SMS
            Set lo = ThisWorkbook.Worksheets(SH_SMS) _
                .ListObjects(TBL_SMS)
        Case Else
            Exit Sub
    End Select
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set changedCells = Intersect(Target, lo.DataBodyRange)
    If changedCells Is Nothing Then Exit Sub

    Set rowsToStamp = CreateObject("Scripting.Dictionary")

    For Each cell In changedCells.Cells
        tableColumnIndex = cell.Column - lo.Range.Column + 1
        If tableColumnIndex >= 1 And _
            tableColumnIndex <= lo.ListColumns.Count Then
            headerName = LCase$(Trim$(lo.ListColumns(tableColumnIndex).Name))
            If headerName <> "current stage" And _
                headerName <> "last updated" And _
                headerName <> "last updated by" Then
                rowsToStamp(CStr(cell.Row)) = cell.Row
            End If
        End If
    Next cell

    currentUser = Trim$(Environ$("Username"))
    If Len(currentUser) = 0 Then currentUser = Trim$(Application.UserName)
    If Len(currentUser) = 0 Then currentUser = "Unknown user"

    For Each rowKey In rowsToStamp.Keys
        tableRowIndex = CLng(rowsToStamp(rowKey)) - lo.DataBodyRange.Row + 1
        If tableRowIndex >= 1 And tableRowIndex <= lo.ListRows.Count Then
            lo.ListColumns("Last Updated").DataBodyRange.Cells( _
                tableRowIndex, 1).Value = Now
            lo.ListColumns("Last Updated").DataBodyRange.Cells( _
                tableRowIndex, 1).NumberFormat = "MM/DD/YYYY h:mm AM/PM"
            lo.ListColumns("Last Updated By").DataBodyRange.Cells( _
                tableRowIndex, 1).Value = currentUser
            lo.ListColumns("Current Stage").DataBodyRange.Cells( _
                tableRowIndex, 1).Calculate
            lo.ListRows(tableRowIndex).Range.WrapText = True
        End If
    Next rowKey

ChangeExit:
End Sub
'''


TOGGLE_INVENTORY_CHECKLIST = r'''
Public Function ToggleInventoryChecklist(ByVal Target As Range) As Boolean
    Dim lo As ListObject
    Dim checklistHeaders As Variant
    Dim item As Variant
    Dim columnKey As String
    Dim columnIndex As Long
    Dim isChecklist As Boolean
    Dim newValue As Boolean

    On Error GoTo ToggleExit

    Select Case Target.Worksheet.Name
        Case SH_EMAIL
            Set lo = ThisWorkbook.Worksheets(SH_EMAIL) _
                .ListObjects(TBL_INVENTORY)
        Case SH_SMS
            Set lo = ThisWorkbook.Worksheets(SH_SMS) _
                .ListObjects(TBL_SMS)
        Case Else
            Exit Function
    End Select
    If lo Is Nothing Then Exit Function
    If Target.CountLarge <> 1 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Function

    columnIndex = Target.Column - lo.Range.Column + 1
    If columnIndex < 1 Or columnIndex > lo.ListColumns.Count Then Exit Function
    columnKey = Trim$(lo.ListColumns(columnIndex).Name)

    checklistHeaders = Array( _
        "Campaign Name and UTM Parameter (Source Code)", _
        "Creative Brief, SL & PH", _
        "SKUs", _
        "In-Design", _
        "Build, QA", _
        "Route", _
        "Send SMS Options", _
        "Send Test", _
        "Approval", _
        "Segments")

    For Each item In checklistHeaders
        If StrComp(columnKey, Trim$(CStr(item)), vbTextCompare) = 0 Then
            isChecklist = True
            Exit For
        End If
    Next item

    If Not isChecklist Then Exit Function

    ' Toggle the checkbox-backed value and update only the affected row.
    If VarType(Target.Value) = vbBoolean Then
        newValue = Not CBool(Target.Value)
        Target.Value = newValue
    ElseIf IsNumeric(Target.Value) Then
        newValue = (CDbl(Target.Value) = 0)
        Target.Value2 = IIf(newValue, 1, 0)
        Target.NumberFormat = LegacyCheckboxNumberFormat()
        Target.Font.Name = "Segoe UI Symbol"
    Else
        newValue = (LCase$(Trim$(CStr(Target.Value))) <> "true")
        Target.Value = newValue
    End If

    ToggleInventoryChecklist = True

ToggleExit:
End Function
'''


STYLE_DASHBOARD = r'''
Private Sub StyleDashboard( _
    ByVal ws As Worksheet, _
    ByVal dashboardTable As ListObject, _
    ByVal summaryRow As Long)

    On Error Resume Next

    ' Style only the active A:L Dashboard workspace.
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

    With ws.Range("A" & summaryRow & ":D" & (summaryRow + 4))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(217, 225, 242)
        .VerticalAlignment = xlCenter
    End With

    ws.Range("A" & summaryRow & ":D" & summaryRow).Interior.Color = _
        RGB(221, 235, 247)
    ws.Range("A" & (summaryRow + 2) & ":D" & (summaryRow + 2)).Interior.Color = _
        RGB(242, 242, 242)
    ws.Range("A" & (summaryRow + 4) & ":D" & (summaryRow + 4)).Interior.Color = _
        RGB(242, 242, 242)

    ws.Columns("A:L").Font.Name = "Aptos"
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 13
    ws.Columns("D").ColumnWidth = 34
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 42
    ws.Columns("F").WrapText = True
    ws.Columns("G").ColumnWidth = 16
    ws.Columns("H:I").ColumnWidth = 12
    ws.Columns("J:L").ColumnWidth = 11

    ApplyDashboardAuditHeader ws
    ApplyDashboardStatusFormatting dashboardTable

    With ws.Range("A1:L" & (summaryRow + 4))
        .WrapText = True
        .VerticalAlignment = xlTop
    End With

    If Not dashboardTable.DataBodyRange Is Nothing Then
        dashboardTable.DataBodyRange.WrapText = True
        dashboardTable.DataBodyRange.VerticalAlignment = xlTop
    End If
End Sub
'''


REBUILD_CALENDARS = r'''
Public Sub RebuildMonthlyCalendars()
    ' Compatibility no-op: monthly calendar sheets are intentionally retired.
End Sub
'''


UPDATE_CALENDAR_TABS = r'''
Public Sub UpdateCalendarTabs()
    ' Compatibility no-op: monthly calendar sheets are intentionally retired.
End Sub
'''


ORDER_WORKBOOK_SHEETS = r'''
Private Sub OrderWorkbookSheets(ByVal wb As Workbook)
    Dim dashboard As Worksheet
    Dim emailSheet As Worksheet
    Dim smsSheet As Worksheet
    Dim notesSheet As Worksheet

    ' Keep the active user-facing sheets in a predictable order.
    Set dashboard = wb.Worksheets(CALENDAR_DASHBOARD_SHEET)
    Set emailSheet = wb.Worksheets(SH_EMAIL)
    Set smsSheet = wb.Worksheets(SH_SMS)
    Set notesSheet = wb.Worksheets(SH_INSTRUCTIONS)

    dashboard.Move Before:=wb.Worksheets(1)
    emailSheet.Move After:=dashboard
    smsSheet.Move After:=emailSheet
    notesSheet.Move After:=smsSheet
End Sub
'''


APPLY_WORKBOOK_WRAP_TEXT = r'''
Private Sub ApplyWorkbookWrapText(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim targetRange As Range

    ' Wrap and auto-fit the active workbook sheets without special calendar logic.
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set targetRange = Nothing

        Select Case ws.Name
            Case SH_DASHBOARD
                Set targetRange = ws.Range("A1:AL180")
            Case SH_INSTRUCTIONS
                Set targetRange = ws.Range("A1:E30")
            Case SH_EMAIL, SH_SMS
                If ws.ListObjects.Count > 0 Then Set targetRange = ws.ListObjects(1).Range
            Case Else
                If Application.WorksheetFunction.CountA(ws.UsedRange) > 0 Then
                    Set targetRange = ws.UsedRange
                End If
        End Select

        If Not targetRange Is Nothing Then
            targetRange.WrapText = True
            targetRange.VerticalAlignment = xlTop
            targetRange.Rows.AutoFit
        End If

        For Each lo In ws.ListObjects
            lo.Range.WrapText = True
            If Not lo.DataBodyRange Is Nothing Then
                lo.DataBodyRange.WrapText = True
                lo.DataBodyRange.VerticalAlignment = xlTop
                lo.DataBodyRange.Rows.AutoFit
            End If
        Next lo

        On Error GoTo 0
    Next ws
End Sub
'''


BUILD_NOTES = r'''
Private Sub BuildNotesInstructionSheet(ByRef wb As Workbook)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nextRow As Long

    ' Rebuild the instructions so configuration refreshes always restore current guidance.
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets(SH_INSTRUCTIONS).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = SH_INSTRUCTIONS
    ws.Tab.Color = RGB(0, 32, 96)

    With ws.Range("A1:E2")
        .Merge
        .Value = "Detailed Notes and Instructions"
        .Font.Name = "Aptos"
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 32, 96)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Range("A3").Value = "Component"
    ws.Range("B3").Value = "Action"
    ws.Range("C3").Value = "How It Works"
    ws.Range("D3").Value = "Required / Avoid"
    ws.Range("E3").Value = "Compatibility Notes"

    nextRow = 4

    AddInstructionRow ws, nextRow, "Email Campaigns", "Add or edit a campaign", _
        "Enter one campaign per table row. Send Date drives the Dashboard's current-week through next-week view.", _
        "Complete Campaign Name and Send Date. Do not rename table headers or delete calculated/audit columns.", _
        "Filters, formulas, and dropdowns work in desktop Excel and Excel for the web."

    AddInstructionRow ws, nextRow, "Email Campaigns", "Track workflow", _
        "Click the workflow cells to toggle checkboxes. Current Stage lists every checked workflow item and shows Completed when all required steps are done.", _
        "Use the checkbox columns as provided. Do not replace them with dropdowns or typed status text.", _
        "Desktop Excel provides click-to-toggle behavior; web users can enter TRUE or FALSE in checkbox-backed cells."

    AddInstructionRow ws, nextRow, "SMS Campaigns", "Add or edit a campaign", _
        "Use the same row-based process as Email Campaigns. SMS workflow uses Send SMS Options, Send Test, Approval, and Segments.", _
        "Complete Campaign Name and Send Date. Keep the SMS table name and headers unchanged.", _
        "The Dashboard combines qualifying Email and SMS rows automatically."

    AddInstructionRow ws, nextRow, "Send Date", "Enter a campaign date", _
        "Enter a real Excel date. It displays as Wednesday, June 10, 2026 while remaining sortable and usable by Dashboard formulas.", _
        "Use a recognized date value. Avoid typing the weekday and month as a plain text sentence.", _
        "The custom dddd, mmmm d, yyyy format is saved in the workbook and works in desktop Excel and Excel for the web."

    AddInstructionRow ws, nextRow, "Send Time", "Enter a time or scheduling label", _
        "Real times display in 12-hour format such as 10:00 AM or 10:00 PM. Text such as STO or Local Timezone is also accepted.", _
        "Do not add blocking validation to this column. Use an Excel time or the required scheduling text.", _
        "Text remains text; numeric Excel times receive the h:mm AM/PM display automatically."

    AddInstructionRow ws, nextRow, "Campaign Type", "Choose or enter a type", _
        "Use Promo, Services, Loyalty & PLCC, Newsletters, Events, NPA, Others, or blank. Custom text is allowed when needed.", _
        "The dropdown source is maintained on the hidden Dropdowns sheet. Do not change its range or delete the sheet.", _
        "Custom entries are not added permanently to the dropdown list."

    AddInstructionRow ws, nextRow, "Filters", "Hide completed or older work", _
        "Use the table filter arrows on Send Date and Delivered to focus on active campaigns or the current month.", _
        "Clear filters before assuming a campaign row is missing.", _
        "Native Excel table filters are SharePoint and web compatible."

    AddInstructionRow ws, nextRow, "Dashboard", "Review upcoming work", _
        "The main table displays Email and SMS campaigns scheduled from the current Sunday through the following Saturday.", _
        "Update source rows rather than typing into Dashboard output cells. Approval shows Done/Not Yet and Segments shows Provided/Pending.", _
        "Native formulas support web viewing; desktop macros provide the full refresh and formatting path."

    AddInstructionRow ws, nextRow, "Dashboard", "Refresh outputs", _
        "Last Refresh is formula-driven. Desktop users can run RefreshDashboard after source-table changes.", _
        "Keep helper columns AA:AL hidden and intact because they support native Dashboard output.", _
        "Excel for the web does not run VBA, but saved formulas and table data remain available."

    AddInstructionRow ws, nextRow, "Retired Features", "Calendars and weekly comparisons", _
        "Monthly Calendar sheets and the Last Week vs Current Week Email/SMS comparison tables and charts were intentionally removed.", _
        "Do not recreate sheets with Calendar in the name or Dashboard comparison objects named DeliveredComparison.", _
        "Legacy macro names remain as safe compatibility wrappers and will not recreate retired features."

    AddInstructionRow ws, nextRow, "Audit Fields", "Last Updated / Last Updated By", _
        "Desktop edits stamp the row timestamp and editor when macros are enabled. Web edits should use SharePoint version history as the authoritative editor record.", _
        "Do not overwrite Last Updated formulas or automation code. Last Updated By remains manually editable for web collaboration.", _
        "Excel for the web cannot execute VBA events or reliably expose the signed-in editor to workbook formulas."

    AddInstructionRow ws, nextRow, "SharePoint", "Collaborate safely", _
        "Store and open the XLSM directly from SharePoint. Coauthoring users should avoid simultaneous structural edits to tables or VBA.", _
        "Do not convert the file to XLSX, because that removes embedded macros. No external Python, PowerShell, or BAS file is required.", _
        "Macros run in desktop Excel only; workbook formulas, tables, filters, and saved results are available in the browser."

    AddInstructionRow ws, nextRow, "Maintenance", "Refresh the Dashboard", _
        "RefreshDashboard recalculates the native current-week through next-week output without rebuilding source tables.", _
        "Create a SharePoint version or backup before intentional table/header redesigns.", _
        "The retired Calendar and comparison features remain removed after configuration refreshes."

    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A3:E" & (nextRow - 1)), , xlYes)
    tbl.Name = "NotesTable"
    tbl.TableStyle = "TableStyleMedium16"

    With ws.Range("A:E")
        .Font.Name = "Aptos"
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 25
    ws.Columns("C").ColumnWidth = 48
    ws.Columns("D").ColumnWidth = 45
    ws.Columns("E").ColumnWidth = 42
    ws.UsedRange.Rows.AutoFit

    ws.Protect Password:="adorama2024", _
        UserInterfaceOnly:=True, AllowFiltering:=True
End Sub
'''


LEGACY_HELPERS = r'''
Public Sub RemoveLegacyCalendarAndComparisonArtifacts(ByVal wb As Workbook)
    Dim sheetIndex As Long
    Dim oldDisplayAlerts As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    oldDisplayAlerts = Application.DisplayAlerts
    On Error GoTo CleanupFailed

    ' Delete every retired Calendar worksheet without prompting.
    Application.DisplayAlerts = False
    For sheetIndex = wb.Worksheets.Count To 1 Step -1
        If IsLegacyCalendarSheetName(wb.Worksheets(sheetIndex).Name) Then
            wb.Worksheets(sheetIndex).Delete
        End If
    Next sheetIndex

    ' Remove retired navigation, tables, charts, values, and formatting.
    RemoveLegacyDashboardComparisons wb.Worksheets(SH_DASHBOARD)

CleanupExit:
    Application.DisplayAlerts = oldDisplayAlerts
    If failureNumber <> 0 Then
        Err.Raise failureNumber, _
            "RemoveLegacyCalendarAndComparisonArtifacts", failureDescription
    End If
    Exit Sub

CleanupFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume CleanupExit
End Sub

Private Sub RemoveLegacyDashboardComparisons(ByVal ws As Worksheet)
    Dim tableIndex As Long
    Dim chartIndex As Long

    On Error Resume Next

    ' Convert retired comparison tables to normal ranges before clearing them.
    For tableIndex = ws.ListObjects.Count To 1 Step -1
        If InStr(1, ws.ListObjects(tableIndex).Name, _
            "DeliveredComparison", vbTextCompare) > 0 Then
            ws.ListObjects(tableIndex).Unlist
        End If
    Next tableIndex

    ' Delete both Email and SMS delivered comparison charts.
    For chartIndex = ws.ChartObjects.Count To 1 Step -1
        If InStr(1, ws.ChartObjects(chartIndex).Name, _
            "Delivered", vbTextCompare) > 0 And _
            InStr(1, ws.ChartObjects(chartIndex).Name, _
            "Comparison", vbTextCompare) > 0 Then
            ws.ChartObjects(chartIndex).Delete
        End If
    Next chartIndex

    ' Clear the retired calendar navigation row and comparison display area.
    ws.Range("A7:L7").Hyperlinks.Delete
    ws.Range("A7:L7").Clear
    ws.Range("N2:S44").UnMerge
    ws.Range("N2:S44").Clear

    On Error GoTo 0
End Sub

Private Function IsLegacyCalendarSheetName(ByVal sheetName As String) As Boolean
    ' Match the former monthly sheets plus any legacy Calendar Import sheet.
    IsLegacyCalendarSheetName = _
        (InStr(1, sheetName, "Calendar", vbTextCompare) > 0)
End Function
'''


THIS_WORKBOOK_OPEN = r'''
Private Sub Workbook_Open()
    Dim priorEvents As Boolean

    priorEvents = Application.EnableEvents
    On Error GoTo SafeExit
    Application.EnableEvents = False

    ' Keep workbook opening light and cross-platform friendly. Excel for the web
    ' does not run VBA; native formulas handle the Dashboard refresh path.
    modEmailProductionTracker.UnfreezeWorkbookViews ThisWorkbook
    ThisWorkbook.Worksheets("Dashboard").Range("B3:D3").Calculate

SafeExit:
    Application.EnableEvents = priorEvents
End Sub
'''


NOTES_ROWS = [
    (
        "Email Campaigns",
        "Add or edit a campaign",
        "Enter one campaign per table row. Send Date drives the Dashboard's current-week through next-week view.",
        "Complete Campaign Name and Send Date. Do not rename table headers or delete calculated/audit columns.",
        "Filters, formulas, and dropdowns work in desktop Excel and Excel for the web.",
    ),
    (
        "Email Campaigns",
        "Track workflow",
        "Click workflow cells to toggle checkboxes. Current Stage lists every checked item and shows Completed when all required steps are done.",
        "Use the checkbox columns as provided. Do not replace them with dropdowns or typed status text.",
        "Desktop Excel provides click-to-toggle behavior; web users can enter TRUE or FALSE in checkbox-backed cells.",
    ),
    (
        "SMS Campaigns",
        "Add or edit a campaign",
        "Use the same row-based process as Email Campaigns. SMS workflow uses Send SMS Options, Send Test, Approval, and Segments.",
        "Complete Campaign Name and Send Date. Keep the SMS table name and headers unchanged.",
        "The Dashboard combines qualifying Email and SMS rows automatically.",
    ),
    (
        "Send Date",
        "Enter a campaign date",
        "Enter a real Excel date. It displays as Wednesday, June 10, 2026 while remaining sortable and usable by Dashboard formulas.",
        "Use a recognized date value. Avoid typing the weekday and month as a plain text sentence.",
        "The custom dddd, mmmm d, yyyy format is saved in the workbook and works in desktop Excel and Excel for the web.",
    ),
    (
        "Send Time",
        "Enter a time or scheduling label",
        "Real times display in 12-hour format such as 10:00 AM or 10:00 PM. Text such as STO or Local Timezone is also accepted.",
        "Do not add blocking validation to this column. Use an Excel time or the required scheduling text.",
        "Text remains text; numeric Excel times receive the h:mm AM/PM display automatically.",
    ),
    (
        "Campaign Type",
        "Choose or enter a type",
        "Use Promo, Services, Loyalty & PLCC, Newsletters, Events, NPA, Others, or blank. Custom text is allowed when needed.",
        "The dropdown source is maintained on the hidden Dropdowns sheet. Do not change its range or delete the sheet.",
        "Custom entries are not added permanently to the dropdown list.",
    ),
    (
        "Filters",
        "Hide completed or older work",
        "Use the table filter arrows on Send Date and Delivered to focus on active campaigns or the current month.",
        "Clear filters before assuming a campaign row is missing.",
        "Native Excel table filters are SharePoint and web compatible.",
    ),
    (
        "Dashboard",
        "Review upcoming work",
        "The main table displays Email and SMS campaigns scheduled from the current Sunday through the following Saturday.",
        "Update source rows rather than typing into Dashboard output cells. Approval shows Done/Not Yet and Segments shows Provided/Pending.",
        "Native formulas support web viewing; desktop macros provide the full refresh and formatting path.",
    ),
    (
        "Dashboard",
        "Refresh outputs",
        "Last Refresh is formula-driven. Desktop users can run RefreshDashboard after source-table changes.",
        "Keep helper columns AA:AL hidden and intact because they support native Dashboard output.",
        "Excel for the web does not run VBA, but saved formulas and table data remain available.",
    ),
    (
        "Retired Features",
        "Calendars and weekly comparisons",
        "Monthly Calendar sheets and Last Week vs Current Week Email/SMS comparison tables and charts were intentionally removed.",
        "Do not recreate sheets with Calendar in the name or Dashboard comparison objects named DeliveredComparison.",
        "Legacy macro names remain as safe compatibility wrappers and will not recreate retired features.",
    ),
    (
        "Audit Fields",
        "Last Updated / Last Updated By",
        "Desktop edits stamp the row timestamp and editor when macros are enabled. Web edits should use SharePoint version history as the authoritative editor record.",
        "Do not overwrite audit formulas or automation code. Last Updated By remains manually editable for web collaboration.",
        "Excel for the web cannot execute VBA events or reliably expose the signed-in editor to workbook formulas.",
    ),
    (
        "SharePoint",
        "Collaborate safely",
        "Store and open the XLSM directly from SharePoint. Avoid simultaneous structural edits to tables or VBA.",
        "Do not convert the file to XLSX. No external Python, PowerShell, or BAS file is required.",
        "Macros run in desktop Excel only; formulas, tables, filters, and saved results are available in the browser.",
    ),
    (
        "Maintenance",
        "Refresh the Dashboard",
        "RefreshDashboard recalculates the native current-week through next-week output without rebuilding source tables.",
        "Create a SharePoint version or backup before intentional table/header redesigns.",
        "Retired Calendar and comparison features remain removed after configuration refreshes.",
    ),
]


def patch_vba(module_text: str) -> str:
    """Update embedded VBA entry points and compatibility wrappers."""
    text = normalize_vba(module_text)

    text = replace_procedure(
        text, "ValidateWorkbookConfiguration", FAST_VALIDATE_WORKBOOK
    )
    text = replace_procedure(
        text, "ApplyAllConfigurations", APPLY_ALL_CONFIGURATIONS
    )
    text = replace_procedure(text, "RefreshDashboard", REFRESH_DASHBOARD_LIGHT)
    text = replace_procedure(text, "ValidateCampaignTable", VALIDATE_CAMPAIGN_TABLE)
    text = replace_procedure(text, "CountBrokenReferences", COUNT_BROKEN_REFERENCES)
    text = replace_procedure(text, "GetInventoryTable", GET_INVENTORY_TABLE)
    text = replace_procedure(text, "GetSmsTable", GET_SMS_TABLE)
    text = replace_procedure(text, "HeaderKey", HEADER_KEY)
    text = replace_procedure(
        text, "IsUserEditableColumn", IS_USER_EDITABLE_COLUMN
    )
    text = replace_procedure(text, "FindTableColumn", FIND_TABLE_COLUMN)
    text = replace_procedure(text, "RefreshNativeOutputs", REFRESH_NATIVE_OUTPUTS)
    text = replace_procedure(
        text, "FormatHyperlinksInChangedCells", FORMAT_CHANGED_HYPERLINKS
    )
    text = replace_procedure(text, "HandleCampaignChange", HANDLE_CAMPAIGN_CHANGE)
    text = replace_procedure(
        text, "ToggleInventoryChecklist", TOGGLE_INVENTORY_CHECKLIST
    )
    text = replace_procedure(text, "StyleDashboard", STYLE_DASHBOARD)
    text = replace_procedure(text, "RebuildMonthlyCalendars", REBUILD_CALENDARS)
    text = replace_procedure(text, "UpdateCalendarTabs", UPDATE_CALENDAR_TABS)
    text = replace_procedure(text, "OrderWorkbookSheets", ORDER_WORKBOOK_SHEETS)
    text = replace_procedure(text, "ApplyWorkbookWrapText", APPLY_WORKBOOK_WRAP_TEXT)
    text = replace_procedure(text, "BuildNotesInstructionSheet", BUILD_NOTES)

    def update_apply_all(proc: str) -> str:
        proc = re.sub(
            r"(?im)^[ \t]*RebuildMonthlyCalendars[ \t]*$",
            "    RemoveLegacyCalendarAndComparisonArtifacts ThisWorkbook",
            proc,
            count=1,
        )
        proc = re.sub(
            r"(?im)^[ \t]*UpdateCalendarTabs[ \t]*$",
            "",
            proc,
            count=1,
        )
        return proc

    if "Public Sub RemoveLegacyCalendarAndComparisonArtifacts" not in text:
        marker = procedure_pattern("CreateDeliveredComparison").search(text)
        if marker is None:
            raise RuntimeError("Could not locate helper insertion point.")
        text = text[: marker.start()] + LEGACY_HELPERS.strip() + "\n\n" + text[marker.start() :]

    text = re.sub(r"\bHeaderKey\s*\(", "NormalizeHeaderKey(", text)
    text = re.sub(
        r"\bNormalizeHeaderKey\s+\(", "NormalizeHeaderKey(", text
    )
    text = re.sub(r"\bFormulaString\s+\(", "FormulaString(", text)
    text = re.sub(r"\.ListObjects\s+\(", ".ListObjects(", text)
    text = re.sub(
        r"(?im)^([ \t]*)Err\.Raise[ \t]+(.+?),[ \t]*,[ \t]*_[ \t]*\n([ \t]*)",
        r"\1Err.Raise Number:=\2, Description:= _\n\3",
        text,
    )

    # Remove excessive blank-line churn before importing the module back into Excel.
    text = re.sub(r"\n{3,}", "\n\n", text)
    return sanitize_vba_continuations(text).replace("\n", "\r\n")


def write_code_module(component, code: str) -> None:
    """Replace one complete VBA code module."""
    code_module = component.CodeModule
    if code_module.CountOfLines:
        retry(
            "delete existing VBA module text",
            lambda: code_module.DeleteLines(1, code_module.CountOfLines),
        )
    retry("import updated VBA module text", lambda: code_module.AddFromString(code))


def remove_legacy_artifacts(workbook) -> None:
    """Remove the retired worksheets, links, tables, charts, and display cells."""
    for sheet_index in range(workbook.Worksheets.Count, 0, -1):
        sheet = workbook.Worksheets(sheet_index)
        if "calendar" in str(sheet.Name).lower():
            sheet.Delete()

    dashboard = workbook.Worksheets(DASHBOARD_SHEET)
    dashboard_was_protected = bool(
        dashboard.ProtectContents
        or dashboard.ProtectDrawingObjects
        or dashboard.ProtectScenarios
    )
    if dashboard_was_protected:
        dashboard.Unprotect(Password="adorama2024")

    for table_index in range(dashboard.ListObjects.Count, 0, -1):
        table = dashboard.ListObjects(table_index)
        if "deliveredcomparison" in str(table.Name).lower():
            table.Unlist()

    for chart_index in range(dashboard.ChartObjects().Count, 0, -1):
        chart = dashboard.ChartObjects().Item(chart_index)
        chart_name = str(chart.Name).lower()
        if "delivered" in chart_name and "comparison" in chart_name:
            try:
                chart.Delete()
            except pywintypes.com_error:
                dashboard.Shapes(chart.Name).Delete()

    try:
        dashboard.Range("A7:L7").Hyperlinks.Delete()
    except pywintypes.com_error:
        pass

    dashboard.Range("A7:L7").Clear()
    dashboard.Range("N2:S44").UnMerge()
    dashboard.Range("N2:S44").Clear()



def rebuild_notes_sheet(workbook) -> None:
    """Create the current Notes - Instructions content directly in the workbook."""
    try:
        workbook.Worksheets(NOTES_SHEET).Delete()
    except pywintypes.com_error:
        pass

    sms_sheet = workbook.Worksheets("SMS Campaigns")
    worksheet = workbook.Worksheets.Add(None, sms_sheet)
    worksheet.Name = NOTES_SHEET
    worksheet.Tab.Color = 6299648

    title = worksheet.Range("A1:E2")
    title.Merge()
    title.Value = "Detailed Notes and Instructions"
    title.Font.Name = "Aptos"
    title.Font.Size = 22
    title.Font.Bold = True
    title.Font.Color = 16777215
    title.Interior.Color = 6299648
    title.HorizontalAlignment = -4108
    title.VerticalAlignment = -4108

    headers = ["Component", "Action", "How It Works", "Required / Avoid", "Compatibility Notes"]
    worksheet.Range("A3:E3").Value = (tuple(headers),)

    first_row = 4
    last_row = first_row + len(NOTES_ROWS) - 1
    worksheet.Range(f"A{first_row}:E{last_row}").Value = tuple(NOTES_ROWS)

    table = worksheet.ListObjects.Add(1, worksheet.Range(f"A3:E{last_row}"), None, 1)
    table.Name = "NotesTable"
    table.TableStyle = "TableStyleMedium16"

    worksheet.Range("A:E").Font.Name = "Aptos"
    worksheet.Range("A:E").VerticalAlignment = -4160
    worksheet.Range("A:E").WrapText = True

    for column, width in zip(("A", "B", "C", "D", "E"), (22, 25, 48, 45, 42)):
        worksheet.Columns(column).ColumnWidth = width

    worksheet.UsedRange.Rows.AutoFit()
    worksheet.Protect(
        Password="adorama2024",
        UserInterfaceOnly=True,
        AllowFiltering=True,
    )


def apply_current_stage_formulas(workbook) -> None:
    """Install formulas that list every checked workflow column."""
    email_formula = (
        '=IF([@[Campaign Name]]="","",LET(checked,TEXTJOIN(", ",TRUE,'
        'IF([@[Campaign Name and UTM Parameter (Source Code)]],'
        '"Campaign Name and UTM Parameter (Source Code)",""),'
        'IF([@[Creative Brief, SL & PH]],"Creative Brief, SL & PH",""),'
        'IF([@SKUs],"SKUs",""),'
        'IF([@[In-Design]],"In-Design",""),'
        'IF([@[Build, QA]],"Build, QA",""),'
        'IF([@Route],"Route",""),'
        'IF([@Approval],"Approval",""),'
        'IF([@Segments],"Segments","")),'
        'IF(checked="","No checklist items checked","Checked: "&checked)))'
    )
    sms_formula = (
        '=IF([@[Campaign Name]]="","",LET(checked,TEXTJOIN(", ",TRUE,'
        'IF([@[Send SMS Options]],"Send SMS Options",""),'
        'IF([@[Send Test]],"Send Test",""),'
        'IF([@Approval],"Approval",""),'
        'IF([@Segments],"Segments","")),'
        'IF(checked="","No checklist items checked","Checked: "&checked)))'
    )

    for sheet_name, table_name, formula in (
        ("Email Campaigns", "EmailCampaignsTable", email_formula),
        ("SMS Campaigns", "SMSCampaignsTable", sms_formula),
    ):
        worksheet = workbook.Worksheets(sheet_name)
        was_protected = bool(
            worksheet.ProtectContents
            or worksheet.ProtectDrawingObjects
            or worksheet.ProtectScenarios
        )
        if was_protected:
            worksheet.Unprotect(Password="adorama2024")
        table = worksheet.ListObjects(table_name)
        stage_column = table.ListColumns("Current Stage")
        stage_column.DataBodyRange.Formula2 = formula
        stage_column.DataBodyRange.WrapText = True


def apply_campaign_formats(workbook) -> None:
    """Normalize date, time, and audit formats across both source tables."""
    for sheet_name, table_name in (
        ("Email Campaigns", "EmailCampaignsTable"),
        ("SMS Campaigns", "SMSCampaignsTable"),
    ):
        worksheet = workbook.Worksheets(sheet_name)
        was_protected = bool(
            worksheet.ProtectContents
            or worksheet.ProtectDrawingObjects
            or worksheet.ProtectScenarios
        )
        if was_protected:
            worksheet.Unprotect(Password="adorama2024")
        table = worksheet.ListObjects(table_name)
        table.ListColumns("Send Date").DataBodyRange.NumberFormat = (
            "dddd, mmmm d, yyyy"
        )
        table.ListColumns("Send Time").DataBodyRange.NumberFormat = "h:mm AM/PM"
        table.ListColumns("Last Updated").DataBodyRange.NumberFormat = (
            "MM/DD/YYYY h:mm AM/PM"
        )
        table.ListColumns("Last Updated By").DataBodyRange.NumberFormat = "@"


def apply_checklist_columns(workbook) -> None:
    """Normalize checklist values and apply native checkbox controls."""
    checked_values = {
        "1",
        "true",
        "yes",
        "x",
        "done",
        "complete",
        "completed",
        chr(0x2611),
    }
    table_columns = (
        (
            "Email Campaigns",
            "EmailCampaignsTable",
            (
                "Campaign Name and UTM Parameter (Source Code)",
                "Creative Brief, SL & PH",
                "SKUs",
                "In-Design",
                "Build, QA",
                "Route",
                "Approval",
                "Segments",
            ),
        ),
        (
            "SMS Campaigns",
            "SMSCampaignsTable",
            ("Send SMS Options", "Send Test", "Approval", "Segments"),
        ),
    )

    for sheet_name, table_name, headers in table_columns:
        worksheet = workbook.Worksheets(sheet_name)
        if (
            worksheet.ProtectContents
            or worksheet.ProtectDrawingObjects
            or worksheet.ProtectScenarios
        ):
            worksheet.Unprotect(Password="adorama2024")

        table = worksheet.ListObjects(table_name)
        for header in headers:
            cell_range = table.ListColumns(header).DataBodyRange
            raw_values = cell_range.Value2
            rows = raw_values if isinstance(raw_values, tuple) else ((raw_values,),)
            normalized = []
            for row in rows:
                value = row[0] if isinstance(row, tuple) else row
                if isinstance(value, bool):
                    normalized_value = value
                elif isinstance(value, (int, float)):
                    normalized_value = value != 0
                else:
                    normalized_value = (
                        str(value or "").strip().lower() in checked_values
                    )
                normalized.append((normalized_value,))

            cell_range.Value2 = tuple(normalized)
            try:
                cell_range.Validation.Delete()
            except pywintypes.com_error:
                pass
            try:
                cell_range.CellControl.SetCheckbox()
            except (AttributeError, pywintypes.com_error):
                unchecked = chr(0x2610)
                checked = chr(0x2611)
                cell_range.NumberFormat = (
                    f'[=1]"{checked}";[=0]"{unchecked}";"{unchecked}"'
                )
                cell_range.Font.Name = "Segoe UI Symbol"
            cell_range.HorizontalAlignment = -4108


def assert_zip_integrity(path: Path) -> None:
    """Verify the edited XLSM package and embedded VBA stream."""
    with zipfile.ZipFile(path) as archive:
        bad_file = archive.testzip()
        if bad_file:
            raise RuntimeError(f"Corrupt workbook package member: {bad_file}")
        if "xl/vbaProject.bin" not in archive.namelist():
            raise RuntimeError("Edited workbook is missing xl/vbaProject.bin.")


def retry(label: str, operation, attempts: int = 20):
    """Retry Excel calls that are temporarily rejected while Excel is busy."""
    last_error = None
    for _ in range(attempts):
        try:
            pythoncom.PumpWaitingMessages()
            return operation()
        except pywintypes.com_error as exc:
            last_error = exc
            if exc.args and exc.args[0] in (-2147418111, -2147417846):
                pythoncom.PumpWaitingMessages()
                continue
            raise
    raise RuntimeError(f"{label} failed repeatedly: {last_error}")


def patch_workbook(source: Path, destination: Path) -> None:
    """Patch one workbook copy through desktop Excel and validate it before closing."""
    print(f"[copy] {source.name}", flush=True)
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)

    print(f"[excel] starting {source.name}", flush=True)
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False
    excel.AutomationSecurity = 1

    workbook = None
    try:
        print(f"[open] {source.name}", flush=True)
        workbook = retry(
            "open workbook",
            lambda: excel.Workbooks.Open(
                str(destination),
                UpdateLinks=0,
                ReadOnly=False,
                AddToMru=False,
            ),
        )

        print(f"[vba] patching {source.name}", flush=True)
        module_component = workbook.VBProject.VBComponents(VBA_MODULE)
        module_code = module_component.CodeModule.Lines(
            1, module_component.CodeModule.CountOfLines
        )
        write_code_module(module_component, patch_vba(module_code))

        workbook_component = workbook.VBProject.VBComponents("ThisWorkbook")
        write_code_module(workbook_component, THIS_WORKBOOK_OPEN.strip().replace("\n", "\r\n"))

        print(f"[sheets] removing retired features from {source.name}", flush=True)
        remove_legacy_artifacts(workbook)
        print(f"[notes] rebuilding instructions in {source.name}", flush=True)
        rebuild_notes_sheet(workbook)
        print(f"[stage] updating Current Stage formulas in {source.name}", flush=True)
        apply_current_stage_formulas(workbook)
        print(f"[format] normalizing campaign formats in {source.name}", flush=True)
        apply_campaign_formats(workbook)
        print(f"[checks] normalizing checklist fields in {source.name}", flush=True)
        apply_checklist_columns(workbook)

        print(f"[save] {source.name}", flush=True)
        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()

    assert_zip_integrity(destination)
    print(f"[done] {destination}", flush=True)


def main() -> int:
    """Patch all named production workbooks into a staging directory."""
    parser = argparse.ArgumentParser()
    parser.add_argument("workbooks", nargs="+", type=Path)
    parser.add_argument("--output-dir", type=Path)
    args = parser.parse_args()

    output_dir = args.output_dir or Path(
        tempfile.mkdtemp(prefix="campaign_tracker_retired_features_")
    )
    output_dir = output_dir.resolve()

    for workbook_path in args.workbooks:
        source = workbook_path.resolve()
        destination = output_dir / source.name
        patch_workbook(source, destination)
        print(destination)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
