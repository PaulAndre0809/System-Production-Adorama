import sys
import pythoncom
import pywintypes
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
        
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0, ReadOnly=True)
        
        out_path = repo_dir / "tools" / "vba_dump.txt"
        with open(out_path, "w", encoding="utf-8") as f:
            for i in range(1, wb.VBProject.VBComponents.Count + 1):
                comp = wb.VBProject.VBComponents(i)
                f.write(f"\n\n--- MODULE: {comp.Name} ---\n")
                count = comp.CodeModule.CountOfLines
                if count > 0:
                    f.write(comp.CodeModule.Lines(1, count))
                
        wb.Close(SaveChanges=False)
        print("Done")
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
