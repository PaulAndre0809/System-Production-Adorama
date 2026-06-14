import win32com.client as win32
from pathlib import Path
import traceback

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    log_path = repo_dir / "rebuild_debug.log"
    if log_path.exists():
        log_path.unlink()
        
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        excel.EnableEvents = True
        
        print("Opening workbook...")
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0)
        
        comp = wb.VBProject.VBComponents("modEmailProductionTracker")
        count = comp.CodeModule.CountOfLines
        code = comp.CodeModule.Lines(1, count)
        
        # We will inject a custom RebuildMonthlyCalendars that writes to rebuild_debug.log
        vba_log_path = str(log_path).replace("\\", "\\\\")
        
        debug_vba = f"""
Public Sub RebuildMonthlyCalendars()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim monthNumber As Long
    Dim fso As Object
    Dim ts As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile("{vba_log_path}", True)
    
    ts.WriteLine "Started RebuildMonthlyCalendars"
    Set wb = ThisWorkbook
    
    On Error GoTo BuildFailed
    
    For monthNumber = 1 To 12
        ts.WriteLine "Month: " & monthNumber & " - GetOrCreate"
        Set ws = GetOrCreateCalendarSheet(wb, monthNumber)
        ts.WriteLine "Month: " & monthNumber & " - BuildCalendarSheet"
        BuildCalendarSheet ws, monthNumber
    Next monthNumber
    
    ts.WriteLine "Deleting README"
    DeleteSheetIfPresent wb, "README"
    ts.WriteLine "Deleting VBA Code"
    DeleteSheetIfPresent wb, "VBA Code"
    
    ts.WriteLine "AddDashboardCalendarLinks"
    AddDashboardCalendarLinks wb
    ts.WriteLine "StyleCoreWorkbookSheets"
    StyleCoreWorkbookSheets wb
    ts.WriteLine "UpdateCalendarTabs"
    UpdateCalendarTabs
    ts.WriteLine "OrderWorkbookSheets"
    OrderWorkbookSheets wb
    ts.WriteLine "ConfigureWorkbookViews"
    ConfigureWorkbookViews wb
    
    ts.WriteLine "Setting dropdowns/log visibility"
    wb.Worksheets("Dropdowns").Visible = xlSheetVeryHidden
    wb.Worksheets("Automation Log").Visible = xlSheetVeryHidden
    
    For monthNumber = 1 To 12
        ts.WriteLine "Calculate Month " & monthNumber
        wb.Worksheets("2026 " & MonthName(monthNumber) & " Calendar").Calculate
    Next monthNumber
    
    ts.WriteLine "Rebuild finished successfully"
    ts.Close
    Exit Sub
    
BuildFailed:
    ts.WriteLine "Failed: " & Err.Description & " (" & Err.Number & ")"
    ts.Close
    Err.Raise Err.Number, "RebuildMonthlyCalendars", Err.Description
End Sub
"""
        # Find where RebuildMonthlyCalendars starts and ends
        start_line = -1
        end_line = -1
        lines = code.splitlines()
        for idx, line in enumerate(lines):
            if "Public Sub RebuildMonthlyCalendars()" in line:
                start_line = idx + 1
            if start_line != -1 and "End Sub" in line and idx + 1 > start_line:
                end_line = idx + 1
                break
                
        if start_line != -1 and end_line != -1:
            print(f"Replacing RebuildMonthlyCalendars lines {start_line} to {end_line}")
            # Delete old sub and insert new one
            comp.CodeModule.DeleteLines(start_line, end_line - start_line + 1)
            comp.CodeModule.InsertLines(start_line, debug_vba)
            print("Running RebuildMonthlyCalendars...")
            try:
                excel.Run(f"'{wb.Name}'!RebuildMonthlyCalendars")
                print("Run complete.")
            except Exception as e:
                print(f"Run failed: {e}")
        else:
            print("RebuildMonthlyCalendars not found in modEmailProductionTracker")
            
        wb.Close(SaveChanges=False)
    except Exception as e:
        print(f"Error occurred: {e}")
        traceback.print_exc()
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass
                
    if log_path.exists():
        print("\n--- Rebuild Debug Log ---")
        print(log_path.read_text())
        print("-------------------------")

if __name__ == "__main__":
    main()
