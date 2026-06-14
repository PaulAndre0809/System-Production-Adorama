import sys
import win32com.client as win32
from pathlib import Path

VBA_CODE = """
' ==============================================================================
' modDynamicHighlighting
' Description: Highlights upcoming scheduled items on the Dashboard.
' ==============================================================================
Option Explicit

Public Sub ApplyDynamicScheduleFilterAndHighlight(ByVal lo As ListObject)
    Dim ws As Worksheet
    Dim dateCol As ListColumn
    Dim dateColIndex As Long
    Dim r As Long
    Dim cellValue As Variant
    Dim cellDate As Date
    Dim todayDate As Date
    Dim currentDayOfWeek As Long
    Dim targetStart As Date
    Dim targetEnd As Date

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Set ws = lo.Parent
    todayDate = Date
    currentDayOfWeek = Weekday(todayDate, vbMonday)

    ' Define target dates
    If currentDayOfWeek >= 1 And currentDayOfWeek <= 4 Then
        targetStart = todayDate + 1
        targetEnd = todayDate + 1
    ElseIf currentDayOfWeek = 5 Then
        targetStart = todayDate + 1
        targetEnd = todayDate + 3
    Else
        targetStart = todayDate + 1
        targetEnd = todayDate + 1
    End If

    On Error Resume Next
    Set dateCol = lo.ListColumns("Send Date")
    On Error GoTo 0

    If dateCol Is Nothing Then Exit Sub
    dateColIndex = dateCol.Index

    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean

    oldCalculation = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    ' Step 1: Clear existing fill
    lo.DataBodyRange.Interior.ColorIndex = xlNone

    ' Step 2: Highlight matching rows visually with a premium soft blue
    For r = 1 To lo.ListRows.Count
        cellValue = lo.DataBodyRange.Cells(r, dateColIndex).Value
        If IsDate(cellValue) Then
            cellDate = DateValue(CDate(cellValue))
            If cellDate >= DateValue(targetStart) And cellDate <= DateValue(targetEnd) Then
                lo.ListRows(r).Range.Interior.Color = RGB(221, 235, 247)
            End If
        End If
    Next r

CleanExit:
    Application.Calculation = oldCalculation
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Exit Sub

ErrorHandler:
    ' Silently fail instead of popping up MsgBoxes that hang python automation
    Resume CleanExit
End Sub
"""

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    template_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker Template.xlsm"
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1
        
        for p in [wb_path, template_path]:
            print(f"Opening workbook: {p.name}")
            wb = excel.Workbooks.Open(str(p), UpdateLinks=0)
            
            try:
                vb_project = wb.VBProject
                
                # Check if modDynamicHighlighting already exists
                mod_name = "modDynamicHighlighting"
                existing_comp = None
                for i in range(1, vb_project.VBComponents.Count + 1):
                    comp = vb_project.VBComponents(i)
                    if comp.Name == mod_name:
                        existing_comp = comp
                        break
                        
                if existing_comp is not None:
                    vb_project.VBComponents.Remove(existing_comp)
                    
                new_comp = vb_project.VBComponents.Add(1)
                new_comp.Name = mod_name
                new_comp.CodeModule.AddFromString(VBA_CODE)
                
                tracker_comp = vb_project.VBComponents("modEmailProductionTracker")
                
                # Check if the call is already there
                found = False
                for i in range(1, tracker_comp.CodeModule.CountOfLines + 1):
                    line = tracker_comp.CodeModule.Lines(i, 1)
                    if "ApplyDynamicScheduleFilterAndHighlight" in line:
                        found = True
                        break
                        
                if not found:
                    for i in range(1, tracker_comp.CodeModule.CountOfLines + 1):
                        line = tracker_comp.CodeModule.Lines(i, 1)
                        if "ApplyDashboardNativeFormulas wsD, dashboardTable" in line:
                            tracker_comp.CodeModule.InsertLines(i + 1, "    Call ApplyDynamicScheduleFilterAndHighlight(dashboardTable)")
                            print(f"Injected Call into {p.name}")
                            break
                            
                wb.Save()
                print(f"Successfully processed {p.name}")
            except Exception as e:
                print(f"Failed to process {p.name}: {e}")
                
            wb.Close(SaveChanges=False)
            
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass

if __name__ == "__main__":
    main()
