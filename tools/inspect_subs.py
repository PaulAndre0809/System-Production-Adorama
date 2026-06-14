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
        
        targets = [1689, 2211, 2227, 4084, 4294, 4568, 4603]
        for t in targets:
            print(f"\n--- Around Line {t} ---")
            start = max(1, t - 5)
            end = min(len(lines), t + 5)
            for idx in range(start, end + 1):
                print(f"{idx}: {lines[idx-1]}")
                
        wb.Close(SaveChanges=False)
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if excel:
            excel.Quit()

if __name__ == "__main__":
    main()
