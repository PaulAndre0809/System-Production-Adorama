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
        ws = wb.Worksheets("Email Campaigns")
        table = ws.ListObjects("EmailCampaignsTable")
        col = None
        for c in table.ListColumns:
            if c.Name == "Campaign Type":
                col = c
                break
                
        if col:
            rng = col.DataBodyRange
            print(f"Cell count: {rng.Count}")
            first_cell = rng.Cells(1, 1)
            try:
                v_type = first_cell.Validation.Type
                formula = first_cell.Validation.Formula1
                show_err = first_cell.Validation.showError
                print(f"Validation Type: {v_type}")
                print(f"Formula: {formula}")
                print(f"Show Error: {show_err}")
            except Exception as e:
                print(f"No validation found or error: {e}")
        else:
            print("Campaign Type column not found")
            
        wb.Close(False)
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if excel:
            excel.Quit()

if __name__ == "__main__":
    main()
