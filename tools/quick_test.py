import sys
import traceback
import win32com.client as win32
from pathlib import Path

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1
        
        print("Opening workbook...")
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0)
        
        macro_name = f"'{wb.Name}'!RefreshProductionStatus"
        print(f"Running macro: {macro_name}")
        
        try:
            excel.Run(macro_name)
            print("Macro ran successfully.")
        except Exception as e:
            print("Error running macro:")
            traceback.print_exc()
            
        print("Closing workbook...")
        wb.Close(SaveChanges=False)
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass

if __name__ == "__main__":
    main()
