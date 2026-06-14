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
        
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0)
        
        comp = wb.VBProject.VBComponents("modEmailProductionTracker")
        count = comp.CodeModule.CountOfLines
        code = comp.CodeModule.Lines(1, count)
        
        lines = code.splitlines()
        current_sub = "Module Level"
        for idx, line in enumerate(lines):
            line_stripped = line.strip()
            if line_stripped.startswith("Public Sub ") or line_stripped.startswith("Private Sub ") or line_stripped.startswith("Public Function ") or line_stripped.startswith("Private Function "):
                current_sub = line_stripped
            
            if "Font.Name" in line_stripped:
                print(f"Line {idx+1} in [{current_sub}]: {line_stripped}")
                
        wb.Close(SaveChanges=False)
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if excel:
            excel.Quit()

if __name__ == "__main__":
    main()
