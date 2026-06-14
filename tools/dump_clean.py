import sys
import win32com.client as win32
from pathlib import Path

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "tools" / "clean_tracker.xlsm"
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3
        
        print(f"Opening {wb_path.name}")
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0, ReadOnly=True)
        
        out_path = repo_dir / "tools" / "clean_vba.txt"
        with open(out_path, "w", encoding="utf-8") as f:
            for i in range(1, wb.VBProject.VBComponents.Count + 1):
                comp = wb.VBProject.VBComponents(i)
                f.write(f"\n' === {comp.Name} ===\n")
                if comp.CodeModule.CountOfLines > 0:
                    code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
                    f.write(code)
        
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
