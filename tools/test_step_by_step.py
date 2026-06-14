import sys
import time
import traceback
import win32com.client as win32
from pathlib import Path

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True # Make Excel visible so we can see what's happening if it blocks
        excel.DisplayAlerts = False
        excel.EnableEvents = True
        
        print("Opening workbook...")
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0)
        
        print("Running: ValidateWorkbookConfiguration...")
        start = time.perf_counter()
        res = excel.Run(f"'{wb.Name}'!ValidateWorkbookConfiguration")
        print(f"ValidateWorkbookConfiguration finished in {time.perf_counter() - start:.2f}s with result: {res}")
        
        print("Running: RebuildMonthlyCalendars...")
        start = time.perf_counter()
        excel.Run(f"'{wb.Name}'!RebuildMonthlyCalendars")
        print(f"RebuildMonthlyCalendars finished in {time.perf_counter() - start:.2f}s")
        
        print("Running: RefreshDashboard...")
        start = time.perf_counter()
        excel.Run(f"'{wb.Name}'!RefreshDashboard")
        print(f"RefreshDashboard finished in {time.perf_counter() - start:.2f}s")
        
        print("Running: UpdateCalendarTabs...")
        start = time.perf_counter()
        excel.Run(f"'{wb.Name}'!UpdateCalendarTabs")
        print(f"UpdateCalendarTabs finished in {time.perf_counter() - start:.2f}s")
        
        print("Running: ApplyAllConfigurations...")
        start = time.perf_counter()
        excel.Run(f"'{wb.Name}'!ApplyAllConfigurations")
        print(f"ApplyAllConfigurations finished in {time.perf_counter() - start:.2f}s")
        
        print("Closing workbook...")
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

if __name__ == "__main__":
    main()
