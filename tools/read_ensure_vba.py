import win32com.client as win32
from pathlib import Path

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.AutomationSecurity = 1
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0)
        comp = wb.VBProject.VBComponents("modEmailProductionTracker")
        code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
        
        idx = code.find("Sub EnsureCampaignSheets()")
        if idx != -1:
            print(code[idx:idx+1500])
        else:
            print("EnsureCampaignSheets not found!")
            
        wb.Close(False)
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if excel:
            excel.Quit()

if __name__ == "__main__":
    main()
