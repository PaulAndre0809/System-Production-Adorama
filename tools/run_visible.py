import win32com.client as win32
from pathlib import Path
import time
import pythoncom
import pywintypes

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True # MAKE IT VISIBLE!
        excel.DisplayAlerts = True # SHOW ALERTS!
        excel.EnableEvents = True # ENABLE EVENTS!
        excel.AutomationSecurity = 1
        
        print("Opening workbook visibly...")
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0)
        
        print("Running RefreshProductionStatus macro...")
        try:
            excel.Run(f"'{wb.Name}'!RefreshProductionStatus")
            print("Macro completed successfully!")
        except Exception as e:
            print(f"Macro failed: {e}")
            
        print("Leaving Excel open for 10 seconds so user can see it...")
        time.sleep(10)
        
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
