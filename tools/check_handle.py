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
        excel.AutomationSecurity = 1
        
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0)
        comp = wb.VBProject.VBComponents("modEmailProductionTracker")
        code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
        
        has_handle = "HandleCampaignChange" in code
        lines_count = comp.CodeModule.CountOfLines
        
        # print first 50 lines and check if it contains the sub
        print(f"Lines count: {lines_count}")
        print(f"Contains HandleCampaignChange: {has_handle}")
        
        # find line containing HandleCampaignChange
        code_lines = code.splitlines()
        for idx, line in enumerate(code_lines):
            if "HandleCampaignChange" in line:
                print(f"Found on line {idx+1}: {line}")
                # Print 10 lines around it
                start = max(0, idx - 5)
                end = min(len(code_lines), idx + 10)
                for j in range(start, end):
                    print(f"{j+1}: {code_lines[j]}")
                break
                
        wb.Close(False)
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if excel:
            excel.Quit()

if __name__ == "__main__":
    main()
