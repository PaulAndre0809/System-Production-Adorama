import sys
import time
from pathlib import Path
import pythoncom
import pywintypes
import win32com.client as win32

def retry(label, func, attempts=20, delay=0.75):
    last_error = None
    for _ in range(attempts):
        try:
            pythoncom.PumpWaitingMessages()
            return func()
        except pywintypes.com_error as exc:
            last_error = exc
            if exc.args and exc.args[0] in (-2147418111, -2147417846):
                time.sleep(delay)
                continue
            raise
    raise RuntimeError(f"{label} failed repeatedly: {last_error}")

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        
        workbook = retry("open", lambda: excel.Workbooks.Open(str(wb_path), UpdateLinks=0, ReadOnly=True))
        sheet = workbook.Worksheets("Notes - Instructions")
        
        for r in range(1, 30):
            row_vals = []
            for c in range(1, 10):
                val = sheet.Cells(r, c).Value
                row_vals.append(str(val) if val is not None else "")
            print(f"Row {r}: {' | '.join(row_vals)}")
            
        workbook.Close(SaveChanges=False)
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass

if __name__ == "__main__":
    main()
