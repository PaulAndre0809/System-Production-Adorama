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

def write_instruction(excel, wb_path):
    print(f"Updating instructions in: {wb_path}")
    workbook = retry("open", lambda: excel.Workbooks.Open(str(wb_path), UpdateLinks=0, ReadOnly=False, AddToMru=False))
    sheet = workbook.Worksheets("Notes - Instructions")
    
    # Find the next empty row
    next_row = 3
    while sheet.Cells(next_row, 1).Value is not None and str(sheet.Cells(next_row, 1).Value).strip() != "":
        next_row += 1
        
    sheet.Cells(next_row, 1).Value = "Dynamic Scheduling Highlight"
    sheet.Cells(next_row, 2).Value = "Automatically applies a visual highlight to campaigns scheduled for the upcoming days depending on the current day of the week."
    sheet.Cells(next_row, 3).Value = "Monday-Thursday: highlights tomorrow's campaigns. Friday-Sunday: highlights campaigns for the upcoming Monday."
    sheet.Cells(next_row, 4).Value = "Works when macros are enabled. Do not remove the modDynamicHighlighting VBA module."
    sheet.Cells(next_row, 5).Value = "Simply check the Dashboard and source sheets; highlighted rows visually indicate upcoming high-priority campaigns."
    
    # Format the row like the others
    # Assuming row 16 has the same formatting, we can copy formats from row 16
    sheet.Range(f"A16:E16").Copy()
    sheet.Range(f"A{next_row}:E{next_row}").PasteSpecial(Paste=-4122) # xlPasteFormats
    excel.Application.CutCopyMode = False
    
    # Apply vertical alignment top and text wrap
    new_range = sheet.Range(f"A{next_row}:E{next_row}")
    new_range.VerticalAlignment = -4160 # xlTop
    new_range.WrapText = True
    
    workbook.Save()
    workbook.Close(SaveChanges=True)

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    
    workbooks = [
        repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm",
        repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker Template.xlsm"
    ]
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        
        for wb in workbooks:
            if wb.exists():
                write_instruction(excel, wb)
            
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
