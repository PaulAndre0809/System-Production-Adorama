"""Apply resilient Send Date and Send Time behavior to tracker workbooks.

This is a development utility. The resulting XLSM files are self-contained and
do not require this script at runtime.
"""

from __future__ import annotations

import argparse
import re
import shutil
import tempfile
from pathlib import Path

import pywintypes
import win32com.client as win32

from retire_calendar_comparisons import (
    VBA_MODULE,
    assert_zip_integrity,
    normalize_vba,
    replace_procedure,
    retry,
    transform_procedure,
    write_code_module,
)


NOTES_SHEET = "Notes - Instructions"
SHEET_PASSWORD = "adorama2024"
SEND_DATE_FORMAT = "dddd, mmmm d, yyyy"
SEND_TIME_FORMAT = "h:mm AM/PM"

EMAIL_ACTIVE_FORMULA = (
    '=SUMPRODUCT(--(EmailCampaignsTable[Campaign Name]<>""),'
    '--(IFERROR(--EmailCampaignsTable[Delivered],0)<=0),'
    '--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>"cancelled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>"canceled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>"cancelled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>"canceled"))'
)
SMS_ACTIVE_FORMULA = (
    '=SUMPRODUCT(--(SMSCampaignsTable[Campaign Name]<>""),'
    '--(IFERROR(--SMSCampaignsTable[Delivered],0)<=0),'
    '--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>"cancelled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>"canceled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>"cancelled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>"canceled"))'
)
SENDING_TODAY_FORMULA = (
    '=SUMPRODUCT(--(IFERROR(INT(EmailCampaignsTable[Send Date]),0)=TODAY()),'
    '--(EmailCampaignsTable[Campaign Name]<>""),'
    '--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>"cancelled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>"canceled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>"cancelled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>"canceled"))'
    '+SUMPRODUCT(--(IFERROR(INT(SMSCampaignsTable[Send Date]),0)=TODAY()),'
    '--(SMSCampaignsTable[Campaign Name]<>""),'
    '--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>"cancelled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>"canceled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>"cancelled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>"canceled"))'
)
APPROVAL_PENDING_FORMULA = (
    '=SUMPRODUCT(--(EmailCampaignsTable[Campaign Name]<>""),'
    '--(EmailCampaignsTable[Approval]=FALSE),'
    '--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>"cancelled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>"canceled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>"cancelled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>"canceled"))'
    '+SUMPRODUCT(--(SMSCampaignsTable[Campaign Name]<>""),'
    '--(SMSCampaignsTable[Approval]=FALSE),'
    '--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>"cancelled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>"canceled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>"cancelled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>"canceled"))'
)
SENT_FORMULA = (
    '=SUMPRODUCT(--(EmailCampaignsTable[Campaign Name]<>""),'
    '--(IFERROR(--EmailCampaignsTable[Delivered],0)>0),'
    '--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>"cancelled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>"canceled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>"cancelled"),'
    '--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>"canceled"))'
    '+SUMPRODUCT(--(SMSCampaignsTable[Campaign Name]<>""),'
    '--(IFERROR(--SMSCampaignsTable[Delivered],0)>0),'
    '--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>"cancelled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>"canceled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>"cancelled"),'
    '--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>"canceled"))'
)

DASHBOARD_SPILL_FORMULA = (
    '=LET(ws,TODAY()-WEEKDAY(TODAY(),1)+1,we,ws+13,'
    'empty,HSTACK("","","","","","","","","","","",""),'
    'eCancel,((LOWER(TRIM(EmailCampaignsTable[Current Stage]))="cancelled")'
    '+(LOWER(TRIM(EmailCampaignsTable[Current Stage]))="canceled")'
    '+(LOWER(TRIM(EmailCampaignsTable[Notes]))="cancelled")'
    '+(LOWER(TRIM(EmailCampaignsTable[Notes]))="canceled"))>0,'
    'email,FILTER(HSTACK(EmailCampaignsTable[Send Date],'
    'EmailCampaignsTable[Send Time],'
    'IF(EmailCampaignsTable[Campaign Name]<>"","Email",""),'
    'EmailCampaignsTable[Campaign Name],EmailCampaignsTable[Campaign Type],'
    'EmailCampaignsTable[Current Stage],EmailCampaignsTable[Owner],'
    'IF(EmailCampaignsTable[Approval],"Done","Not Yet"),'
    'IF(EmailCampaignsTable[Segments],"Provided","Pending"),'
    'EmailCampaignsTable[Jira Link],EmailCampaignsTable[ClickUp Link],'
    'EmailCampaignsTable[Bluecore/Attentive Link]),'
    '(EmailCampaignsTable[Campaign Name]<>"")'
    '*(IFERROR(INT(EmailCampaignsTable[Send Date]),0)>=ws)'
    '*(IFERROR(INT(EmailCampaignsTable[Send Date]),0)<=we)'
    '*(eCancel=FALSE),empty),'
    'sCancel,((LOWER(TRIM(SMSCampaignsTable[Current Stage]))="cancelled")'
    '+(LOWER(TRIM(SMSCampaignsTable[Current Stage]))="canceled")'
    '+(LOWER(TRIM(SMSCampaignsTable[Notes]))="cancelled")'
    '+(LOWER(TRIM(SMSCampaignsTable[Notes]))="canceled"))>0,'
    'sms,FILTER(HSTACK(SMSCampaignsTable[Send Date],'
    'SMSCampaignsTable[Send Time],'
    'IF(SMSCampaignsTable[Campaign Name]<>"","SMS",""),'
    'SMSCampaignsTable[Campaign Name],SMSCampaignsTable[Campaign Type],'
    'SMSCampaignsTable[Current Stage],SMSCampaignsTable[Owner],'
    'IF(SMSCampaignsTable[Approval],"Done","Not Yet"),'
    'IF(SMSCampaignsTable[Segments],"Provided","Pending"),'
    'SMSCampaignsTable[Proof of Schedule],"",'
    'SMSCampaignsTable[Bluecore/Attentive Link]),'
    '(SMSCampaignsTable[Campaign Name]<>"")'
    '*(IFERROR(INT(SMSCampaignsTable[Send Date]),0)>=ws)'
    '*(IFERROR(INT(SMSCampaignsTable[Send Date]),0)<=we)'
    '*(sCancel=FALSE),empty),'
    'data,VSTACK(email,sms),'
    'clean,FILTER(data,CHOOSECOLS(data,4)<>""),'
    'IFERROR(SORTBY(clean,CHOOSECOLS(clean,1),1,'
    'IFERROR(--CHOOSECOLS(clean,2),0),1,'
    'CHOOSECOLS(clean,3),1,CHOOSECOLS(clean,4),1),""))'
)


FORMAT_SEND_DATE_COLUMN = r'''
Private Sub FormatSendDateColumn(ByVal lo As ListObject)
    Dim lc As ListColumn

    Set lc = FindTableColumn(lo, "Send Date")
    If lc Is Nothing Then
        Err.Raise vbObjectError + 1030, , _
            "Send Date column not found."
    End If

    If Not lc.DataBodyRange Is Nothing Then
        ' Store real Excel dates while using a portable long-date display.
        lc.DataBodyRange.NumberFormat = "dddd, mmmm d, yyyy"
        lc.DataBodyRange.WrapText = True
    End If

    lc.Range.ColumnWidth = 28
End Sub
'''


FORMAT_SEND_TIME_COLUMN = r'''
Private Sub FormatSendTimeColumn(ByVal lo As ListObject)
    Dim lc As ListColumn

    Set lc = FindTableColumn(lo, "Send Time")
    If lc Is Nothing Then Exit Sub

    If Not lc.DataBodyRange Is Nothing Then
        ' Do not use blocking validation: time values and labels such as STO or
        ' Local Timezone must both remain valid campaign inputs.
        On Error Resume Next
        lc.DataBodyRange.Validation.Delete
        On Error GoTo 0

        lc.DataBodyRange.NumberFormat = "h:mm AM/PM"
        lc.DataBodyRange.WrapText = True
    End If

    lc.Range.ColumnWidth = 20
End Sub
'''


APPLY_CAMPAIGN_ENTRY_FORMATS = r'''
Public Sub ApplyCampaignEntryFormats()
    Dim emailTable As ListObject
    Dim smsTable As ListObject
    Dim dashboardTable As ListObject
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo FormatFailed

    Set emailTable = ThisWorkbook.Worksheets(SH_EMAIL) _
        .ListObjects(TBL_INVENTORY)
    Set smsTable = ThisWorkbook.Worksheets(SH_SMS) _
        .ListObjects(TBL_SMS)

    FormatSendDateColumn emailTable
    FormatSendTimeColumn emailTable
    FormatSendDateColumn smsTable
    FormatSendTimeColumn smsTable

    ' Keep the Dashboard presentation consistent with the source tables.
    Set dashboardTable = ThisWorkbook.Worksheets(SH_DASHBOARD) _
        .ListObjects(TBL_DASHBOARD)
    If Not dashboardTable.DataBodyRange Is Nothing Then
        dashboardTable.ListColumns("Send Date").DataBodyRange _
            .NumberFormat = "dddd, mmmm d, yyyy"
        dashboardTable.ListColumns("Time").DataBodyRange _
            .NumberFormat = "h:mm AM/PM"
        dashboardTable.ListColumns("Send Date").DataBodyRange.WrapText = True
        dashboardTable.ListColumns("Time").DataBodyRange.WrapText = True
    End If
    dashboardTable.ListColumns("Send Date").Range.ColumnWidth = 28
    dashboardTable.ListColumns("Time").Range.ColumnWidth = 24

    Exit Sub

FormatFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    LogAction "ApplyCampaignEntryFormats", _
        "Error " & CStr(failureNumber) & ": " & failureDescription
    On Error GoTo 0
End Sub
'''


APPLY_DASHBOARD_KPI_FORMULAS = r'''
Public Sub ApplyDashboardKpiFormulas()
    Dim wsD As Worksheet

    On Error GoTo KpiExit
    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)

    ' Full table-column references avoid implicit-intersection errors and
    ' automatically expand when campaign rows are added.
    ApplyFormulaCompat wsD.Range("E5"), _
        "=SUMPRODUCT(--(EmailCampaignsTable[Campaign Name]<>"""")," & _
        "--(IFERROR(--EmailCampaignsTable[Delivered],0)<=0)," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>""cancelled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>""canceled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>""cancelled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>""canceled""))"

    ApplyFormulaCompat wsD.Range("G5"), _
        "=SUMPRODUCT(--(SMSCampaignsTable[Campaign Name]<>"""")," & _
        "--(IFERROR(--SMSCampaignsTable[Delivered],0)<=0)," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>""cancelled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>""canceled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>""cancelled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>""canceled""))"

    ApplyFormulaCompat wsD.Range("A5"), "=E5+G5"

    ApplyFormulaCompat wsD.Range("C5"), _
        "=SUMPRODUCT(--(IFERROR(INT(EmailCampaignsTable[Send Date]),0)=TODAY())," & _
        "--(EmailCampaignsTable[Campaign Name]<>"""")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>""cancelled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>""canceled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>""cancelled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>""canceled""))+" & _
        "SUMPRODUCT(--(IFERROR(INT(SMSCampaignsTable[Send Date]),0)=TODAY())," & _
        "--(SMSCampaignsTable[Campaign Name]<>"""")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>""cancelled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>""canceled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>""cancelled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>""canceled""))"

    ApplyFormulaCompat wsD.Range("I5"), _
        "=SUMPRODUCT(--(EmailCampaignsTable[Campaign Name]<>"""")," & _
        "--(EmailCampaignsTable[Approval]=FALSE)," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>""cancelled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>""canceled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>""cancelled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>""canceled""))+" & _
        "SUMPRODUCT(--(SMSCampaignsTable[Campaign Name]<>"""")," & _
        "--(SMSCampaignsTable[Approval]=FALSE)," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>""cancelled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>""canceled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>""cancelled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>""canceled""))"

    ApplyFormulaCompat wsD.Range("K5"), _
        "=SUMPRODUCT(--(EmailCampaignsTable[Campaign Name]<>"""")," & _
        "--(IFERROR(--EmailCampaignsTable[Delivered],0)>0)," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>""cancelled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Current Stage]))<>""canceled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>""cancelled"")," & _
        "--(LOWER(TRIM(EmailCampaignsTable[Notes]))<>""canceled""))+" & _
        "SUMPRODUCT(--(SMSCampaignsTable[Campaign Name]<>"""")," & _
        "--(IFERROR(--SMSCampaignsTable[Delivered],0)>0)," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>""cancelled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Current Stage]))<>""canceled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>""cancelled"")," & _
        "--(LOWER(TRIM(SMSCampaignsTable[Notes]))<>""canceled""))"

    ' Keep the legacy footer accurate without maintaining duplicate logic.
    ApplyFormulaCompat wsD.Range("C162"), "=C5"
    ApplyFormulaCompat wsD.Range("C163"), "=E5"
    ApplyFormulaCompat wsD.Range("C164"), "=G5"
    ApplyFormulaCompat wsD.Range("C165"), "=I5"
    ApplyFormulaCompat wsD.Range("C166"), "=K5"

KpiExit:
End Sub
'''


FORMAT_TIMED_HYPERLINKS = r'''
Private Sub FormatHyperlinksInChangedCells( _
    ByVal lo As ListObject, _
    ByVal changedCells As Range)

    Dim cell As Range
    Dim tableColumnIndex As Long
    Dim headerName As String
    Dim displayName As String

    On Error GoTo LinkExit

    For Each cell In changedCells.Cells
        tableColumnIndex = cell.Column - lo.Range.Column + 1
        If tableColumnIndex >= 1 And _
            tableColumnIndex <= lo.ListColumns.Count Then

            headerName = CStr(lo.ListColumns(tableColumnIndex).Name)
            displayName = CampaignLinkDisplayName(headerName)

            If Len(displayName) > 0 Then
                InstallTimedCampaignLink lo, cell, displayName
            End If
        End If
    Next cell

LinkExit:
End Sub
'''


TIMED_HYPERLINK_HELPERS = r'''
Public Sub ApplyTimedCampaignLinks()
    Dim emailTable As ListObject
    Dim smsTable As ListObject

    On Error GoTo LinkExit

    Set emailTable = ThisWorkbook.Worksheets(SH_EMAIL) _
        .ListObjects(TBL_INVENTORY)
    Set smsTable = ThisWorkbook.Worksheets(SH_SMS) _
        .ListObjects(TBL_SMS)

    ApplyTimedLinksToTable emailTable, Array( _
        "Jira Link", "ClickUp Link", "Bluecore/Attentive Link")
    ApplyTimedLinksToTable smsTable, Array( _
        "Proof of Schedule", "Bluecore/Attentive Link")

LinkExit:
End Sub

Private Sub ApplyTimedLinksToTable( _
    ByVal lo As ListObject, _
    ByVal linkHeaders As Variant)

    Dim item As Variant
    Dim linkColumn As ListColumn
    Dim cell As Range
    Dim displayName As String

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For Each item In linkHeaders
        Set linkColumn = FindTableColumn(lo, CStr(item))
        If Not linkColumn Is Nothing Then
            displayName = CampaignLinkDisplayName(CStr(item))
            For Each cell In linkColumn.DataBodyRange.Cells
                InstallTimedCampaignLink lo, cell, displayName
            Next cell
        End If
    Next item
End Sub

Private Function CampaignLinkDisplayName( _
    ByVal headerName As String) As String

    Select Case LCase$(Trim$(headerName))
        Case "jira link"
            CampaignLinkDisplayName = "JIRA"
        Case "clickup link"
            CampaignLinkDisplayName = "ClickUp"
        Case "bluecore/attentive link", "bluecore link"
            CampaignLinkDisplayName = "Bluecore/Attentive"
        Case "proof of schedule"
            CampaignLinkDisplayName = "Proof of Schedule"
    End Select
End Function

Private Function CampaignLinkAddress(ByVal cell As Range) As String
    Dim formulaText As String
    Dim firstQuote As Long
    Dim quotePosition As Long
    Dim rawValue As String

    On Error Resume Next
    If cell.Hyperlinks.Count > 0 Then
        CampaignLinkAddress = Trim$(CStr(cell.Hyperlinks(1).Address))
    End If
    On Error GoTo 0

    If Len(CampaignLinkAddress) = 0 And Not cell.HasFormula Then
        rawValue = Trim$(CStr(cell.Value2))
        If LCase$(Left$(rawValue, 4)) = "http" Then
            CampaignLinkAddress = rawValue
        End If
    End If

    If Len(CampaignLinkAddress) = 0 And cell.HasFormula Then
        formulaText = CStr(cell.Formula)
        If InStr(1, formulaText, "=HYPERLINK(", vbTextCompare) = 1 Then
            firstQuote = InStr(1, formulaText, Chr$(34))
            quotePosition = firstQuote + 1
            Do While quotePosition > firstQuote
                quotePosition = InStr(quotePosition, formulaText, Chr$(34))
                If quotePosition = 0 Then Exit Do
                If Mid$(formulaText, quotePosition + 1, 1) = Chr$(34) Then
                    quotePosition = quotePosition + 2
                Else
                    CampaignLinkAddress = Mid$( _
                        formulaText, firstQuote + 1, _
                        quotePosition - firstQuote - 1)
                    CampaignLinkAddress = Replace( _
                        CampaignLinkAddress, Chr$(34) & Chr$(34), Chr$(34))
                    Exit Do
                End If
            Loop
        End If
    End If
End Function

Private Sub InstallTimedCampaignLink( _
    ByVal lo As ListObject, _
    ByVal cell As Range, _
    ByVal displayName As String)

    Dim linkAddress As String
    Dim formulaText As String

    linkAddress = CampaignLinkAddress(cell)
    If LCase$(Left$(linkAddress, 4)) <> "http" Then Exit Sub

    formulaText = TimedCampaignLinkFormula( _
        linkAddress, displayName)

    If CStr(cell.Formula) <> formulaText Then
        ApplyFormulaCompat cell, formulaText
    End If
End Sub

Private Function TimedCampaignLinkFormula( _
    ByVal linkAddress As String, _
    ByVal displayName As String) As String

    Dim q As String
    Dim safeAddress As String
    Dim safeDisplay As String
    Dim sendDateRef As String
    Dim sendTimeRef As String
    Dim maturityFormula As String

    q = Chr$(34)
    safeAddress = Replace(linkAddress, q, q & q)
    safeDisplay = Replace(displayName, q, q & q)
    sendDateRef = "[@[Send Date]]"
    sendTimeRef = "[@[Send Time]]"

    ' Numeric times use the exact timestamp. Text scheduling labels such as
    ' STO or Local Timezone use midnight on Send Date as the fallback.
    maturityFormula = _
        "INT(" & sendDateRef & ")+IF(ISNUMBER(" & _
        sendTimeRef & "),MOD(" & sendTimeRef & ",1),0)+7"

    TimedCampaignLinkFormula = _
        "=HYPERLINK(" & q & safeAddress & q & "," & _
        "IF(AND(ISNUMBER(" & sendDateRef & "),NOW()>=" & _
        maturityFormula & ")," & q & safeDisplay & q & "," & _
        q & safeAddress & q & "))"
End Function

Public Sub RefreshTimedCampaignLinks()
    CalculateTimedCampaignLinkColumns
    ScheduleNextCampaignLinkRefresh
End Sub

Private Sub CalculateTimedCampaignLinkColumns()
    Dim emailTable As ListObject
    Dim smsTable As ListObject

    On Error GoTo CalculateExit
    Set emailTable = ThisWorkbook.Worksheets(SH_EMAIL) _
        .ListObjects(TBL_INVENTORY)
    Set smsTable = ThisWorkbook.Worksheets(SH_SMS) _
        .ListObjects(TBL_SMS)

    emailTable.ListColumns("Jira Link").DataBodyRange.Calculate
    emailTable.ListColumns("ClickUp Link").DataBodyRange.Calculate
    emailTable.ListColumns("Bluecore/Attentive Link") _
        .DataBodyRange.Calculate
    smsTable.ListColumns("Proof of Schedule").DataBodyRange.Calculate
    smsTable.ListColumns("Bluecore/Attentive Link") _
        .DataBodyRange.Calculate

CalculateExit:
End Sub

Public Sub ScheduleNextCampaignLinkRefresh()
    Dim nextRefresh As Date

    On Error GoTo ScheduleExit
    CancelCampaignLinkRefresh
    nextRefresh = NextCampaignLinkMaturity()

    If nextRefresh > Now Then
        mNextCampaignLinkRefresh = nextRefresh
        Application.OnTime _
            EarliestTime:=mNextCampaignLinkRefresh, _
            Procedure:="'" & ThisWorkbook.Name & _
                "'!RefreshTimedCampaignLinks", _
            Schedule:=True
    End If

ScheduleExit:
End Sub

Public Sub CancelCampaignLinkRefresh()
    On Error Resume Next
    If mNextCampaignLinkRefresh > 0 Then
        Application.OnTime _
            EarliestTime:=mNextCampaignLinkRefresh, _
            Procedure:="'" & ThisWorkbook.Name & _
                "'!RefreshTimedCampaignLinks", _
            Schedule:=False
    End If
    mNextCampaignLinkRefresh = 0
    On Error GoTo 0
End Sub

Private Function NextCampaignLinkMaturity() As Date
    Dim candidate As Date

    candidate = NextTableLinkMaturity( _
        ThisWorkbook.Worksheets(SH_EMAIL).ListObjects(TBL_INVENTORY), _
        Array("Jira Link", "ClickUp Link", _
            "Bluecore/Attentive Link"))
    If candidate > Now Then NextCampaignLinkMaturity = candidate

    candidate = NextTableLinkMaturity( _
        ThisWorkbook.Worksheets(SH_SMS).ListObjects(TBL_SMS), _
        Array("Proof of Schedule", "Bluecore/Attentive Link"))
    If candidate > Now Then
        If NextCampaignLinkMaturity = 0 Or _
            candidate < NextCampaignLinkMaturity Then
            NextCampaignLinkMaturity = candidate
        End If
    End If
End Function

Private Function NextTableLinkMaturity( _
    ByVal lo As ListObject, _
    ByVal linkHeaders As Variant) As Date

    Dim rowIndex As Long
    Dim item As Variant
    Dim linkColumn As ListColumn
    Dim hasLink As Boolean
    Dim maturityValue As Date

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To lo.ListRows.Count
        hasLink = False
        For Each item In linkHeaders
            Set linkColumn = FindTableColumn(lo, CStr(item))
            If Not linkColumn Is Nothing Then
                If Len(CampaignLinkAddress( _
                    linkColumn.DataBodyRange.Cells(rowIndex, 1))) > 0 Then
                    hasLink = True
                    Exit For
                End If
            End If
        Next item

        If hasLink Then
            maturityValue = CampaignLinkMaturity( _
                lo, rowIndex)
            If maturityValue > Now Then
                If NextTableLinkMaturity = 0 Or _
                    maturityValue < NextTableLinkMaturity Then
                    NextTableLinkMaturity = maturityValue
                End If
            End If
        End If
    Next rowIndex
End Function

Private Function CampaignLinkMaturity( _
    ByVal lo As ListObject, _
    ByVal rowIndex As Long) As Date

    Dim sendDateValue As Variant
    Dim sendTimeValue As Variant
    Dim baseDate As Double
    Dim timePart As Double

    sendDateValue = lo.ListColumns("Send Date") _
        .DataBodyRange.Cells(rowIndex, 1).Value2
    sendTimeValue = lo.ListColumns("Send Time") _
        .DataBodyRange.Cells(rowIndex, 1).Value2

    If Not IsNumeric(sendDateValue) Then Exit Function
    baseDate = Int(CDbl(sendDateValue))

    If IsNumeric(sendTimeValue) Then
        timePart = CDbl(sendTimeValue) - _
            Int(CDbl(sendTimeValue))
    End If

    CampaignLinkMaturity = CDate(baseDate + timePart + 7)
End Function
'''


DASHBOARD_NATIVE_SPILL_FORMULA = r'''
Private Function DashboardNativeSpillFormula() As String
    Dim q As String
    Dim emptyRow As String
    Dim emailRows As String
    Dim smsRows As String
    Dim emailCancelled As String
    Dim smsCancelled As String

    q = Chr$(34)
    emptyRow = DashboardNativeBlankRowFormula()

    emailCancelled = _
        "((LOWER(TRIM(EmailCampaignsTable[Current Stage]))=" & _
        q & "cancelled" & q & ")+" & _
        "(LOWER(TRIM(EmailCampaignsTable[Current Stage]))=" & _
        q & "canceled" & q & ")+" & _
        "(LOWER(TRIM(EmailCampaignsTable[Notes]))=" & _
        q & "cancelled" & q & ")+" & _
        "(LOWER(TRIM(EmailCampaignsTable[Notes]))=" & _
        q & "canceled" & q & "))=0"

    smsCancelled = _
        "((LOWER(TRIM(SMSCampaignsTable[Current Stage]))=" & _
        q & "cancelled" & q & ")+" & _
        "(LOWER(TRIM(SMSCampaignsTable[Current Stage]))=" & _
        q & "canceled" & q & ")+" & _
        "(LOWER(TRIM(SMSCampaignsTable[Notes]))=" & _
        q & "cancelled" & q & ")+" & _
        "(LOWER(TRIM(SMSCampaignsTable[Notes]))=" & _
        q & "canceled" & q & "))=0"

    emailRows = _
        "FILTER(HSTACK(" & _
        "EmailCampaignsTable[Send Date]," & _
        "EmailCampaignsTable[Send Time]," & _
        "IF(EmailCampaignsTable[Campaign Name]<>" & q & q & "," & _
            q & "Email" & q & "," & q & q & ")," & _
        "EmailCampaignsTable[Campaign Name]," & _
        "EmailCampaignsTable[Campaign Type]," & _
        "EmailCampaignsTable[Current Stage]," & _
        "EmailCampaignsTable[Owner]," & _
        "IF(EmailCampaignsTable[Approval]," & q & "Done" & q & "," & _
            q & "Not Yet" & q & ")," & _
        "IF(EmailCampaignsTable[Segments]," & q & "Provided" & q & "," & _
            q & "Pending" & q & ")," & _
        "EmailCampaignsTable[Jira Link]," & _
        "EmailCampaignsTable[ClickUp Link]," & _
        "EmailCampaignsTable[Bluecore/Attentive Link])," & _
        "(EmailCampaignsTable[Campaign Name]<>" & q & q & ")*" & _
        "(IFERROR(INT(EmailCampaignsTable[Send Date]),0)>=ws)*" & _
        "(IFERROR(INT(EmailCampaignsTable[Send Date]),0)<=we)*" & _
        "(" & emailCancelled & "),empty)"

    smsRows = _
        "FILTER(HSTACK(" & _
        "SMSCampaignsTable[Send Date]," & _
        "SMSCampaignsTable[Send Time]," & _
        "IF(SMSCampaignsTable[Campaign Name]<>" & q & q & "," & _
            q & "SMS" & q & "," & q & q & ")," & _
        "SMSCampaignsTable[Campaign Name]," & _
        "SMSCampaignsTable[Campaign Type]," & _
        "SMSCampaignsTable[Current Stage]," & _
        "SMSCampaignsTable[Owner]," & _
        "IF(SMSCampaignsTable[Approval]," & q & "Done" & q & "," & _
            q & "Not Yet" & q & ")," & _
        "IF(SMSCampaignsTable[Segments]," & q & "Provided" & q & "," & _
            q & "Pending" & q & ")," & _
        "SMSCampaignsTable[Proof of Schedule]," & _
        q & q & "," & _
        "SMSCampaignsTable[Bluecore/Attentive Link])," & _
        "(SMSCampaignsTable[Campaign Name]<>" & q & q & ")*" & _
        "(IFERROR(INT(SMSCampaignsTable[Send Date]),0)>=ws)*" & _
        "(IFERROR(INT(SMSCampaignsTable[Send Date]),0)<=we)*" & _
        "(" & smsCancelled & "),empty)"

    DashboardNativeSpillFormula = _
        "=LET(ws,TODAY()-WEEKDAY(TODAY(),1)+1," & _
        "we,ws+13," & _
        "empty," & emptyRow & "," & _
        "email," & emailRows & "," & _
        "sms," & smsRows & "," & _
        "data,VSTACK(email,sms)," & _
        "clean,FILTER(data,CHOOSECOLS(data,4)<>" & q & q & ")," & _
        "IFERROR(SORTBY(clean,CHOOSECOLS(clean,1),1," & _
        "IFERROR(--CHOOSECOLS(clean,2),0),1," & _
        "CHOOSECOLS(clean,3),1,CHOOSECOLS(clean,4),1)," & _
        q & q & "))"
End Function
'''


REFRESH_DASHBOARD = r'''
Public Sub RefreshDashboard()
    Dim wsD As Worksheet
    Dim dashboardTable As ListObject

    On Error GoTo RefreshExit

    Set wsD = ThisWorkbook.Worksheets(SH_DASHBOARD)
    Set dashboardTable = wsD.ListObjects(TBL_DASHBOARD)

    ApplyDashboardKpiFormulas
    CalculateTimedCampaignLinkColumns
    ApplyFormulaCompat wsD.Range("AA11"), _
        DashboardNativeSpillFormula()
    wsD.Range("AA11").Calculate

    If Not dashboardTable.DataBodyRange Is Nothing Then
        dashboardTable.DataBodyRange.Calculate
    End If

    wsD.Range("A5:K5").Calculate
    wsD.Range("B3:D3").Calculate
    ScheduleNextCampaignLinkRefresh

RefreshExit:
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
    Dim updatedColumn As ListColumn
    Dim updatedByColumn As ListColumn
    Dim stageColumn As ListColumn
    Dim stageCells As Range
    Dim formatRows As Range
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo ChangeFailed

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

    ' Restore display formats after typing or pasting into date/time cells.
    For Each cell In changedCells.Cells
        tableColumnIndex = cell.Column - lo.Range.Column + 1
        If tableColumnIndex >= 1 And _
            tableColumnIndex <= lo.ListColumns.Count Then

            headerName = LCase$(Trim$( _
                lo.ListColumns(tableColumnIndex).Name))

            Select Case headerName
                Case "send date"
                    cell.NumberFormat = "dddd, mmmm d, yyyy"
                    cell.WrapText = True
                Case "send time"
                    On Error Resume Next
                    cell.Validation.Delete
                    On Error GoTo ChangeFailed
                    cell.NumberFormat = "h:mm AM/PM"
                    cell.WrapText = True
            End Select
        End If
    Next cell

    ' Preserve friendly hyperlinks when users paste raw URLs.
    FormatHyperlinksInChangedCells lo, changedCells

    Set rowsToStamp = CreateObject("Scripting.Dictionary")

    For Each cell In changedCells.Cells
        tableColumnIndex = cell.Column - lo.Range.Column + 1
        If tableColumnIndex >= 1 And _
            tableColumnIndex <= lo.ListColumns.Count Then
            headerName = LCase$(Trim$( _
                lo.ListColumns(tableColumnIndex).Name))
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

    Set updatedColumn = lo.ListColumns("Last Updated")
    Set updatedByColumn = lo.ListColumns("Last Updated By")
    Set stageColumn = lo.ListColumns("Current Stage")

    For Each rowKey In rowsToStamp.Keys
        tableRowIndex = CLng(rowsToStamp(rowKey)) - _
            lo.DataBodyRange.Row + 1
        If tableRowIndex >= 1 And _
            tableRowIndex <= lo.ListRows.Count Then
            updatedColumn.DataBodyRange.Cells( _
                tableRowIndex, 1).Value = Now
            updatedColumn.DataBodyRange.Cells( _
                tableRowIndex, 1).NumberFormat = _
                "MM/DD/YYYY h:mm AM/PM"
            updatedByColumn.DataBodyRange.Cells( _
                tableRowIndex, 1).Value = currentUser

            If stageCells Is Nothing Then
                Set stageCells = stageColumn.DataBodyRange.Cells( _
                    tableRowIndex, 1)
                Set formatRows = lo.ListRows(tableRowIndex).Range
            Else
                Set stageCells = Union(stageCells, _
                    stageColumn.DataBodyRange.Cells(tableRowIndex, 1))
                Set formatRows = Union(formatRows, _
                    lo.ListRows(tableRowIndex).Range)
            End If
        End If
    Next rowKey

    If Not stageCells Is Nothing Then stageCells.Calculate
    If Not formatRows Is Nothing Then
        formatRows.WrapText = True
        formatRows.Rows.AutoFit
    End If

    ' Recalculate once per edit batch, not once per changed cell.
    If rowsToStamp.Count > 0 Then RefreshDashboard
    Exit Sub

ChangeFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    LogAction "HandleCampaignChange", _
        ws.Name & " error " & CStr(failureNumber) & ": " & _
        failureDescription
    On Error GoTo 0
End Sub
'''


WORKBOOK_OPEN = r'''
Private Sub Workbook_Open()
    Dim priorEvents As Boolean

    priorEvents = Application.EnableEvents
    On Error GoTo SafeExit
    Application.EnableEvents = False

    ' Keep workbook opening light. Saved formats work in Excel for the web;
    ' desktop Excel also repairs date/time formatting when the file opens.
    modEmailProductionTracker.UnfreezeWorkbookViews ThisWorkbook
    modEmailProductionTracker.ApplyCampaignEntryFormats
    modEmailProductionTracker.ApplyTimedCampaignLinks
    modEmailProductionTracker.RefreshDashboard

SafeExit:
    Application.EnableEvents = priorEvents
End Sub
'''


WORKBOOK_BEFORE_CLOSE = r'''
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modEmailProductionTracker.CancelCampaignLinkRefresh
    On Error GoTo 0
End Sub
'''


VBA_NOTES_ROWS = r'''
    AddInstructionRow ws, nextRow, "Send Date", "Enter a campaign date", _
        "Enter a real Excel date. It displays as Wednesday, June 10, 2026 while remaining sortable and usable by Dashboard formulas.", _
        "Use a recognized date value. Avoid typing the weekday and month as a plain text sentence.", _
        "The custom dddd, mmmm d, yyyy format is saved in the workbook and works in desktop Excel and Excel for the web."

    AddInstructionRow ws, nextRow, "Send Time", "Enter a time or scheduling label", _
        "Real times display in 12-hour format such as 10:00 AM or 10:00 PM. Text such as STO or Local Timezone is also accepted.", _
        "Do not add blocking validation to this column. Use an Excel time or the required scheduling text.", _
        "Text remains text; numeric Excel times receive the h:mm AM/PM display automatically."
'''

VBA_MAINTENANCE_NOTES_ROWS = r'''
    AddInstructionRow ws, nextRow, "Cancelled Campaigns", "Exclude from Dashboard", _
        "Rows are excluded when Current Stage or Notes is exactly Cancelled or Canceled, ignoring capitalization and surrounding spaces.", _
        "Use the exact cancellation status. Do not expect partial phrases such as Cancelled by team to be treated as the status.", _
        "The Dashboard campaign list and summary KPIs use the same cancellation rule."

    AddInstructionRow ws, nextRow, "Timed Link Labels", "Convert URLs after seven days", _
        "JIRA, ClickUp, Bluecore/Attentive, and Proof of Schedule links display the full URL until seven days after Send Date and Send Time, then display the platform name.", _
        "Use a real Send Date and numeric Excel time for an exact timestamp. STO, Local Timezone, or blank time uses midnight on Send Date.", _
        "Native HYPERLINK and NOW formulas work in desktop Excel and Excel for the web. Desktop Excel also schedules the next due refresh while open."

    AddInstructionRow ws, nextRow, "Troubleshooting", "VBA compile and edit errors", _
        "The embedded VBA project is compiled during QA. Line-continuation characters must be followed immediately by the next VBA line.", _
        "Do not insert blank lines after an underscore continuation character. Keep macros enabled in desktop Excel for event automation.", _
        "Excel for the web does not compile or run VBA; saved formulas and formats continue to work there."

    AddInstructionRow ws, nextRow, "Performance", "Keep calculations responsive", _
        "Dashboard KPIs and the current-week feed use expanding table references. Desktop edits recalculate once per edit batch.", _
        "Do not replace table formulas with entire-column worksheet references or rename tables and headers.", _
        "RefreshDashboard repairs KPI formulas and recalculates only the Dashboard output ranges."

    AddInstructionRow ws, nextRow, "Sheet Protection", "Maintain the instruction sheet", _
        "Notes - Instructions is protected to prevent accidental structural edits.", _
        "Administrative password: adorama2024. Reprotect the sheet after intentional maintenance.", _
        "Worksheet protection is an editing safeguard, not encryption or access control."
'''


NOTES_ROWS = (
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
)

MAINTENANCE_NOTES_ROWS = (
    (
        "Cancelled Campaigns",
        "Exclude from Dashboard",
        "Rows are excluded when Current Stage or Notes is exactly Cancelled or Canceled, ignoring capitalization and surrounding spaces.",
        "Use the exact cancellation status. Do not expect partial phrases such as Cancelled by team to be treated as the status.",
        "The Dashboard campaign list and summary KPIs use the same cancellation rule.",
    ),
    (
        "Timed Link Labels",
        "Convert URLs after seven days",
        "JIRA, ClickUp, Bluecore/Attentive, and Proof of Schedule links display the full URL until seven days after Send Date and Send Time, then display the platform name.",
        "Use a real Send Date and numeric Excel time for an exact timestamp. STO, Local Timezone, or blank time uses midnight on Send Date.",
        "Native HYPERLINK and NOW formulas work in desktop Excel and Excel for the web. Desktop Excel also schedules the next due refresh while open.",
    ),
    (
        "Troubleshooting",
        "VBA compile and edit errors",
        "The embedded VBA project is compiled during QA. Line-continuation characters must be followed immediately by the next VBA line.",
        "Do not insert blank lines after an underscore continuation character. Keep macros enabled in desktop Excel for event automation.",
        "Excel for the web does not compile or run VBA; saved formulas and formats continue to work there.",
    ),
    (
        "Performance",
        "Keep calculations responsive",
        "Dashboard KPIs and the current-week feed use expanding table references. Desktop edits recalculate once per edit batch.",
        "Do not replace table formulas with entire-column worksheet references or rename tables and headers.",
        "RefreshDashboard repairs KPI formulas and recalculates only the Dashboard output ranges.",
    ),
    (
        "Sheet Protection",
        "Maintain the instruction sheet",
        "Notes - Instructions is protected to prevent accidental structural edits.",
        "Administrative password: adorama2024. Reprotect the sheet after intentional maintenance.",
        "Worksheet protection is an editing safeguard, not encryption or access control.",
    ),
)


def patch_notes_builder(procedure: str) -> str:
    """Add the current date/time guidance to the embedded Notes builder."""
    updated = procedure

    if "Wednesday, June 10, 2026" not in updated:
        marker = re.search(
            r'(?im)^[ \t]*AddInstructionRow ws, nextRow, "Campaign Type"',
            updated,
        )
        if marker is None:
            raise RuntimeError("Could not locate the Notes builder insertion point.")
        updated = (
            updated[: marker.start()]
            + VBA_NOTES_ROWS.strip()
            + "\n\n"
            + updated[marker.start() :]
        )

    if '"Cancelled Campaigns", "Exclude from Dashboard"' not in updated:
        marker = re.search(
            r"(?im)^[ \t]*Set[ \t]+tbl[ \t]*=",
            updated,
        )
        if marker is None:
            raise RuntimeError("Could not locate the Notes table creation point.")
        updated = (
            updated[: marker.start()]
            + VBA_MAINTENANCE_NOTES_ROWS.strip()
            + "\n\n"
            + updated[marker.start() :]
        )

    return updated


def patch_vba(module_text: str) -> str:
    """Patch only the entry-formatting and edit-event procedures."""
    text = normalize_vba(module_text)
    text = replace_procedure(
        text, "FormatSendDateColumn", FORMAT_SEND_DATE_COLUMN
    )
    text = replace_procedure(
        text, "FormatSendTimeColumn", FORMAT_SEND_TIME_COLUMN
    )
    text = replace_procedure(
        text, "HandleCampaignChange", HANDLE_CAMPAIGN_CHANGE
    )
    text = replace_procedure(
        text,
        "FormatHyperlinksInChangedCells",
        FORMAT_TIMED_HYPERLINKS,
    )
    text = replace_procedure(
        text,
        "DashboardNativeSpillFormula",
        DASHBOARD_NATIVE_SPILL_FORMULA,
    )
    text = replace_procedure(text, "RefreshDashboard", REFRESH_DASHBOARD)

    if "mNextCampaignLinkRefresh" not in text:
        marker = re.search(
            r"(?im)^[ \t]*(?:Public|Private)[ \t]+Const\b",
            text,
        )
        if marker is None:
            raise RuntimeError("Could not locate VBA constant insertion point.")
        text = (
            text[: marker.start()]
            + "Private mNextCampaignLinkRefresh As Date\n\n"
            + text[marker.start() :]
        )

    if re.search(
        r"(?im)^[ \t]*Public[ \t]+Sub[ \t]+ApplyCampaignEntryFormats\b",
        text,
    ):
        text = replace_procedure(
            text, "ApplyCampaignEntryFormats", APPLY_CAMPAIGN_ENTRY_FORMATS
        )
    else:
        marker = re.search(
            r"(?im)^[ \t]*Private[ \t]+Sub[ \t]+FormatSendDateColumn\b",
            text,
        )
        if marker is None:
            raise RuntimeError("Could not locate FormatSendDateColumn.")
        text = (
            text[: marker.start()]
            + APPLY_CAMPAIGN_ENTRY_FORMATS.strip()
            + "\n\n"
            + text[marker.start() :]
        )

    if re.search(
        r"(?im)^[ \t]*Public[ \t]+Sub[ \t]+ApplyTimedCampaignLinks\b",
        text,
    ):
        timed_procedures = (
            "ApplyTimedCampaignLinks",
            "ApplyTimedLinksToTable",
            "CampaignLinkDisplayName",
            "CampaignLinkAddress",
            "InstallTimedCampaignLink",
            "TimedCampaignLinkFormula",
            "RefreshTimedCampaignLinks",
            "CalculateTimedCampaignLinkColumns",
            "ScheduleNextCampaignLinkRefresh",
            "CancelCampaignLinkRefresh",
            "NextCampaignLinkMaturity",
            "NextTableLinkMaturity",
            "CampaignLinkMaturity",
        )
        replacement_text = TIMED_HYPERLINK_HELPERS
        for procedure_name in timed_procedures:
            replacement = procedure_pattern_from_block(
                replacement_text,
                procedure_name,
            )
            text = replace_procedure(text, procedure_name, replacement)
    else:
        marker = re.search(
            r"(?im)^[ \t]*Public[ \t]+Sub[ \t]+RefreshDashboard\b",
            text,
        )
        if marker is None:
            raise RuntimeError("Could not locate timed-link insertion point.")
        text = (
            text[: marker.start()]
            + TIMED_HYPERLINK_HELPERS.strip()
            + "\n\n"
            + text[marker.start() :]
        )

    text = re.sub(
        r'DashboardNativeLinkFormula\(([ \t]*)10,[ \t]*"Jira"\)',
        r'DashboardNativeLinkFormula(\g<1>10, "JIRA")',
        text,
        flags=re.IGNORECASE,
    )

    if re.search(
        r"(?im)^[ \t]*Public[ \t]+Sub[ \t]+ApplyDashboardKpiFormulas\b",
        text,
    ):
        text = replace_procedure(
            text,
            "ApplyDashboardKpiFormulas",
            APPLY_DASHBOARD_KPI_FORMULAS,
        )
    else:
        marker = re.search(
            r"(?im)^[ \t]*Public[ \t]+Sub[ \t]+RefreshDashboard\b",
            text,
        )
        if marker is None:
            raise RuntimeError("Could not locate RefreshDashboard.")
        text = (
            text[: marker.start()]
            + APPLY_DASHBOARD_KPI_FORMULAS.strip()
            + "\n\n"
            + text[marker.start() :]
        )

    text = transform_procedure(
        text, "BuildNotesInstructionSheet", patch_notes_builder
    )
    text = re.sub(r"\n{3,}", "\n\n", text)
    return sanitize_vba_code(text).replace("\n", "\r\n")


def sanitize_vba_code(code: str) -> str:
    """Remove illegal blank lines after VBA continuation characters."""
    lines = normalize_vba(code).split("\n")
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


def assert_valid_vba_continuations(code: str, component_name: str) -> None:
    """Fail if any VBA continuation is separated from its next line."""
    lines = normalize_vba(code).split("\n")
    malformed = [
        index + 1
        for index, line in enumerate(lines[:-1])
        if line.rstrip().endswith("_") and not lines[index + 1].strip()
    ]
    if malformed:
        raise RuntimeError(
            f"{component_name} contains malformed VBA continuations: "
            f"{malformed[:20]}"
        )


def update_notes_sheet(workbook) -> None:
    """Update the existing protected Notes table without rebuilding the sheet."""
    worksheet = workbook.Worksheets(NOTES_SHEET)
    worksheet.Unprotect(Password=SHEET_PASSWORD)
    table = worksheet.ListObjects("NotesTable")

    maintained_components = {
        "Send Date",
        "Send Time",
        "Cancelled Campaigns",
        "Timed Link Labels",
        "Troubleshooting",
        "Performance",
        "Sheet Protection",
    }
    for index in range(table.ListRows.Count, 0, -1):
        component = str(table.ListRows(index).Range.Cells(1, 1).Value or "")
        if component in maintained_components:
            table.ListRows(index).Delete()

    for values in (*NOTES_ROWS, *MAINTENANCE_NOTES_ROWS):
        row = table.ListRows.Add()
        row.Range.Value = (values,)

    worksheet.Range("A:E").WrapText = True
    worksheet.UsedRange.Rows.AutoFit()
    worksheet.Protect(
        Password=SHEET_PASSWORD,
        UserInterfaceOnly=True,
        AllowFiltering=True,
    )


def apply_physical_formats(workbook) -> None:
    """Apply saved formats while leaving all campaign values untouched."""
    for sheet_name, table_name in (
        ("Email Campaigns", "EmailCampaignsTable"),
        ("SMS Campaigns", "SMSCampaignsTable"),
    ):
        worksheet = workbook.Worksheets(sheet_name)
        if worksheet.ProtectContents:
            worksheet.Unprotect(Password=SHEET_PASSWORD)

        table = worksheet.ListObjects(table_name)
        send_date = table.ListColumns("Send Date")
        send_time = table.ListColumns("Send Time")

        send_date.DataBodyRange.NumberFormat = SEND_DATE_FORMAT
        send_date.DataBodyRange.WrapText = True
        send_date.Range.ColumnWidth = 28

        try:
            send_time.DataBodyRange.Validation.Delete()
        except pywintypes.com_error:
            pass
        send_time.DataBodyRange.NumberFormat = SEND_TIME_FORMAT
        send_time.DataBodyRange.WrapText = True
        send_time.Range.ColumnWidth = 20

    dashboard = workbook.Worksheets("Dashboard")
    dashboard_table = dashboard.ListObjects("DashboardWorkTable")
    dashboard_date = dashboard_table.ListColumns("Send Date")
    dashboard_time = dashboard_table.ListColumns("Time")
    dashboard_date.DataBodyRange.NumberFormat = SEND_DATE_FORMAT
    dashboard_time.DataBodyRange.NumberFormat = SEND_TIME_FORMAT
    dashboard_date.DataBodyRange.WrapText = True
    dashboard_time.DataBodyRange.WrapText = True
    dashboard_date.Range.ColumnWidth = 28
    dashboard_time.Range.ColumnWidth = 24


def apply_dashboard_kpi_formulas(workbook) -> None:
    """Install accurate, expanding Dashboard KPI formulas."""
    dashboard = workbook.Worksheets("Dashboard")
    formulas = {
        "A5": "=E5+G5",
        "C5": SENDING_TODAY_FORMULA,
        "E5": EMAIL_ACTIVE_FORMULA,
        "G5": SMS_ACTIVE_FORMULA,
        "I5": APPROVAL_PENDING_FORMULA,
        "K5": SENT_FORMULA,
    }
    for address, formula in formulas.items():
        dashboard.Range(address).Formula2 = formula
    dashboard.Range("AA11").Formula2 = DASHBOARD_SPILL_FORMULA
    for footer_address, source_address in {
        "C162": "C5",
        "C163": "E5",
        "C164": "G5",
        "C165": "I5",
        "C166": "K5",
    }.items():
        dashboard.Range(footer_address).Formula2 = f"={source_address}"
    dashboard.Range("A5:K5").Calculate()


def procedure_pattern_from_block(block: str, name: str) -> str:
    """Extract one VBA procedure from a block containing several procedures."""
    normalized = normalize_vba(block)
    pattern = re.compile(
        r"(?ims)^[ \t]*(?:Public|Private|Friend)?[ \t]+"
        r"(?:Sub|Function)[ \t]+"
        + re.escape(name)
        + r"\b.*?^[ \t]*End (?:Sub|Function)[ \t]*$"
    )
    match = pattern.search(normalized)
    if match is None:
        raise RuntimeError(f"Procedure {name} is missing from helper block.")
    return match.group(0)


def hyperlink_address(cell) -> str:
    """Read the underlying URL from a plain or formula hyperlink cell."""
    try:
        if cell.Hyperlinks.Count:
            address = str(cell.Hyperlinks(1).Address or "").strip()
            if address.lower().startswith("http"):
                return address
    except pywintypes.com_error:
        pass

    if not cell.HasFormula:
        value = str(cell.Value2 or "").strip()
        if value.lower().startswith("http"):
            return value

    formula = str(cell.Formula2 or "")
    match = re.match(
        r'^=HYPERLINK\("((?:""|[^"])*)"',
        formula,
        flags=re.IGNORECASE,
    )
    if match:
        return match.group(1).replace('""', '"')
    return ""


def timed_hyperlink_formula(address: str, display_name: str) -> str:
    """Build the native seven-day conditional hyperlink formula."""
    safe_address = address.replace('"', '""')
    safe_display = display_name.replace('"', '""')
    maturity = (
        "INT([@[Send Date]])+"
        "IF(ISNUMBER([@[Send Time]]),MOD([@[Send Time]],1),0)+7"
    )
    return (
        f'=HYPERLINK("{safe_address}",'
        f'IF(AND(ISNUMBER([@[Send Date]]),NOW()>={maturity}),'
        f'"{safe_display}","{safe_address}"))'
    )


def apply_timed_hyperlink_formulas(workbook) -> None:
    """Normalize all supported source links to native timed formulas."""
    configurations = (
        (
            "Email Campaigns",
            "EmailCampaignsTable",
            (
                ("Jira Link", "JIRA"),
                ("ClickUp Link", "ClickUp"),
                ("Bluecore/Attentive Link", "Bluecore/Attentive"),
            ),
        ),
        (
            "SMS Campaigns",
            "SMSCampaignsTable",
            (
                ("Proof of Schedule", "Proof of Schedule"),
                ("Bluecore/Attentive Link", "Bluecore/Attentive"),
            ),
        ),
    )

    for sheet_name, table_name, columns in configurations:
        table = workbook.Worksheets(sheet_name).ListObjects(table_name)
        for header, display_name in columns:
            link_range = table.ListColumns(header).DataBodyRange
            for row_index in range(1, link_range.Rows.Count + 1):
                cell = link_range.Cells(row_index, 1)
                address = hyperlink_address(cell)
                if not address:
                    continue
                formula = timed_hyperlink_formula(address, display_name)
                if str(cell.Formula2) != formula:
                    cell.Formula2 = formula

    dashboard_table = workbook.Worksheets("Dashboard").ListObjects(
        "DashboardWorkTable"
    )
    jira_column = dashboard_table.ListColumns("Jira").DataBodyRange
    for row_index in range(1, jira_column.Rows.Count + 1):
        cell = jira_column.Cells(row_index, 1)
        formula = str(cell.Formula2 or "")
        if '"Jira"' in formula:
            cell.Formula2 = formula.replace('"Jira"', '"JIRA"')


def trim_stale_used_ranges(workbook) -> None:
    """Clear only empty formatting below known live workbook regions."""
    email_sheet = workbook.Worksheets("Email Campaigns")
    email_table = email_sheet.ListObjects("EmailCampaignsTable")
    email_last_row = email_table.Range.Row + email_table.Range.Rows.Count - 1
    email_used_last_row = (
        email_sheet.UsedRange.Row + email_sheet.UsedRange.Rows.Count - 1
    )
    if email_used_last_row > email_last_row:
        candidate = email_sheet.Range(
            f"A{email_last_row + 1}:X{email_used_last_row}"
        )
        if workbook.Application.WorksheetFunction.CountA(candidate) == 0:
            candidate.Clear()

    dashboard = workbook.Worksheets("Dashboard")
    dashboard_live_last_row = 166
    dashboard_used_last_row = (
        dashboard.UsedRange.Row + dashboard.UsedRange.Rows.Count - 1
    )
    if dashboard_used_last_row > dashboard_live_last_row:
        candidate = dashboard.Range(
            f"A{dashboard_live_last_row + 1}:AL{dashboard_used_last_row}"
        )
        if workbook.Application.WorksheetFunction.CountA(candidate) == 0:
            candidate.Clear()

    # Accessing UsedRange after clearing asks Excel to recalculate its boundary.
    _ = email_sheet.UsedRange.Address
    _ = dashboard.UsedRange.Address


def verify_notes_password(workbook) -> None:
    """Prove that the expected password unlocks and relocks the Notes sheet."""
    worksheet = workbook.Worksheets(NOTES_SHEET)
    worksheet.Unprotect(Password=SHEET_PASSWORD)
    if worksheet.ProtectContents:
        raise RuntimeError("Notes sheet password did not remove protection.")
    worksheet.Protect(
        Password=SHEET_PASSWORD,
        UserInterfaceOnly=True,
        AllowFiltering=True,
    )
    if not worksheet.ProtectContents:
        raise RuntimeError("Notes sheet did not relock after password test.")


def patch_workbook(source: Path, destination: Path) -> None:
    """Patch one workbook copy and preserve its existing campaign data."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False
    excel.AutomationSecurity = 1

    workbook = None
    try:
        workbook = retry(
            "open workbook",
            lambda: excel.Workbooks.Open(
                str(destination),
                UpdateLinks=0,
                ReadOnly=False,
                AddToMru=False,
            ),
        )

        module_component = workbook.VBProject.VBComponents(VBA_MODULE)
        module_code = module_component.CodeModule.Lines(
            1, module_component.CodeModule.CountOfLines
        )
        write_code_module(module_component, patch_vba(module_code))

        workbook_component = workbook.VBProject.VBComponents("ThisWorkbook")
        workbook_code = workbook_component.CodeModule.Lines(
            1, workbook_component.CodeModule.CountOfLines
        )
        workbook_code = replace_procedure(
            normalize_vba(workbook_code),
            "Workbook_Open",
            WORKBOOK_OPEN,
        )
        if re.search(
            r"(?im)^[ \t]*Private[ \t]+Sub[ \t]+Workbook_BeforeClose\b",
            workbook_code,
        ):
            workbook_code = replace_procedure(
                workbook_code,
                "Workbook_BeforeClose",
                WORKBOOK_BEFORE_CLOSE,
            )
        else:
            workbook_code = (
                workbook_code.rstrip()
                + "\n\n"
                + WORKBOOK_BEFORE_CLOSE.strip()
                + "\n"
            )
        workbook_code = workbook_code.replace("\n", "\r\n")
        write_code_module(workbook_component, workbook_code)

        for component_index in range(
            1,
            workbook.VBProject.VBComponents.Count + 1,
        ):
            component = workbook.VBProject.VBComponents(component_index)
            code_module = component.CodeModule
            if not code_module.CountOfLines:
                continue
            component_code = code_module.Lines(1, code_module.CountOfLines)
            sanitized_code = sanitize_vba_code(component_code)
            assert_valid_vba_continuations(
                sanitized_code,
                str(component.Name),
            )
            normalized_component = normalize_vba(component_code)
            if sanitized_code != normalized_component:
                write_code_module(
                    component,
                    sanitized_code.replace("\n", "\r\n"),
                )

        apply_physical_formats(workbook)
        apply_dashboard_kpi_formulas(workbook)
        apply_timed_hyperlink_formulas(workbook)
        trim_stale_used_ranges(workbook)
        update_notes_sheet(workbook)
        verify_notes_password(workbook)

        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()

    assert_zip_integrity(destination)


def main() -> int:
    """Patch all supplied workbooks into an isolated staging directory."""
    parser = argparse.ArgumentParser()
    parser.add_argument("workbooks", nargs="+", type=Path)
    parser.add_argument("--output-dir", type=Path)
    args = parser.parse_args()

    output_dir = (
        args.output_dir.resolve()
        if args.output_dir
        else Path(tempfile.mkdtemp(prefix="campaign_entry_formats_"))
    )

    for workbook_path in args.workbooks:
        source = workbook_path.resolve()
        destination = output_dir / source.name
        patch_workbook(source, destination)
        print(destination)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
