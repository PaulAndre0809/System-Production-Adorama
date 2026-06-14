import sys
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
        
        print("Opening workbook...")
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0, ReadOnly=True)
        ws = wb.Worksheets("Automation Log")
        lo = ws.ListObjects("AutomationLogTable")
        
        if lo.ListRows.Count == 0:
            print("No log entries found.")
        else:
            print(f"Total log rows: {lo.ListRows.Count}")
            print("Last 15 log entries:")
            print(f"{'Timestamp':<22} | {'User':<15} | {'Action':<25} | Details")
            print("-" * 100)
            
            # Read in bulk
            vals = lo.DataBodyRange.Value
            if not isinstance(vals, tuple):
                vals = ((vals,),)
            
            # Print last 15 rows
            start_row = max(0, len(vals) - 15)
            for i in range(start_row, len(vals)):
                row = vals[i]
                timestamp = str(row[0]) if len(row) > 0 else ""
                user = str(row[1]) if len(row) > 1 else ""
                action = str(row[2]) if len(row) > 2 else ""
                details = str(row[3]) if len(row) > 3 else ""
                print(f"{timestamp:<22} | {user:<15} | {action:<25} | {details}")
                
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
