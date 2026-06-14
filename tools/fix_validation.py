import win32com.client as win32
from pathlib import Path
import time
import pythoncom
import pywintypes

def retry(func, attempts=10, delay=0.5):
    for i in range(attempts):
        try:
            pythoncom.PumpWaitingMessages()
            return func()
        except pywintypes.com_error as e:
            if i == attempts - 1:
                raise
            time.sleep(delay)

def fix_validation_in_wb(excel, path):
    print(f"Fixing Campaign Type validation in {path.name}")
    wb = retry(lambda: excel.Workbooks.Open(str(path), UpdateLinks=0))
    
    try:
        sheets_and_tables = [
            ("Email Campaigns", "EmailCampaignsTable"),
            ("SMS Campaigns", "SMSCampaignsTable")
        ]
        
        for ws_name, tbl_name in sheets_and_tables:
            ws = wb.Worksheets(ws_name)
            
            # Unprotect sheet
            try:
                ws.Unprotect()
            except Exception as e:
                print(f"  Warning: Unprotect failed on {ws_name}: {e}")
                
            table = ws.ListObjects(tbl_name)
            col = None
            for c in table.ListColumns:
                if c.Name == "Campaign Type":
                    col = c
                    break
                    
            if col is not None:
                rng = col.DataBodyRange
                print(f"  Updating validation for {ws_name} -> {tbl_name} ({rng.Count} cells)")
                
                # We can set validation on the whole range at once
                val = rng.Validation
                val.Delete()
                val.Add(Type=3, AlertStyle=1, Operator=1, Formula1="=Dropdowns!$A$2:$A$9")
                val.IgnoreBlank = True
                val.InCellDropdown = True
                val.ShowInput = True
                val.ShowError = False
            else:
                print(f"  Error: 'Campaign Type' column not found in {tbl_name}")
                
            # Re-protect sheet (mimic original protection if needed, or just protect)
            try:
                ws.Protect()
            except Exception as e:
                print(f"  Warning: Protect failed on {ws_name}: {e}")
                
        wb.Save()
        print(f"Successfully fixed validation in {path.name}")
    except Exception as e:
        print(f"Failed to process {path.name}: {e}")
    finally:
        wb.Close(SaveChanges=False)

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    template_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker Template.xlsm"
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1
        
        for p in [wb_path, template_path]:
            fix_validation_in_wb(excel, p)
            
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
