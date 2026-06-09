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
        
        for i in range(1, wb.VBProject.VBComponents.Count + 1):
            comp = wb.VBProject.VBComponents(i)
            if comp.Name == "modDynamicHighlighting":
                count = comp.CodeModule.CountOfLines
                if count > 0:
                    print(comp.CodeModule.Lines(1, count))
                
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
