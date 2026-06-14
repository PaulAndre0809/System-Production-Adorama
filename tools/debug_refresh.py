import win32com.client as win32
from pathlib import Path
import time

def run_step(excel, workbook, step_name):
    start = time.perf_counter()
    print(f"Running: {step_name}...", end="", flush=True)
    try:
        excel.Run(f"'{workbook.Name}'!{step_name}")
        print(f" Success ({time.perf_counter() - start:.2f}s)")
    except Exception as e:
        print(f" Failed ({time.perf_counter() - start:.2f}s): {e}")

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
        
        print("Opening workbook...")
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0)
        
        # Disable calculation and screen updating to match RefreshProductionStatus VBA
        excel.Calculation = -4135 # xlCalculationManual
        excel.ScreenUpdating = False
        
        run_step(excel, wb, "EnsureCampaignSheets")
        
        # Note: EnsureRenamedColumn and EnsureNotesColumn take parameters, but let's see if we can get ListObjects
        try:
            email_table = wb.Worksheets("Email Campaigns").ListObjects("EmailCampaignsTable")
            sms_table = wb.Worksheets("SMS Campaigns").ListObjects("SMSCampaignsTable")
            
            print("Running EnsureRenamedColumn/EnsureNotesColumn/ApplyCalculatedColumns...")
            start = time.perf_counter()
            excel.Run(f"'{wb.Name}'!EnsureRenamedColumn", email_table, "Bluecore Link", "Bluecore/Attentive Link")
            excel.Run(f"'{wb.Name}'!EnsureRenamedColumn", sms_table, "Bluecore Link", "Bluecore/Attentive Link")
            excel.Run(f"'{wb.Name}'!EnsureNotesColumn", email_table)
            excel.Run(f"'{wb.Name}'!EnsureNotesColumn", sms_table)
            excel.Run(f"'{wb.Name}'!ApplyCalculatedColumns", email_table)
            excel.Run(f"'{wb.Name}'!ApplyCalculatedColumns", sms_table)
            print(f"  Completed calculated columns ({time.perf_counter() - start:.2f}s)")
        except Exception as e:
            print(f"  Failed calculated columns step: {e}")
            
        run_step(excel, wb, "RefreshDashboard")
        
        print("Enabling calculation and screen updating...")
        excel.Calculation = -4105 # xlCalculationAutomatic
        excel.ScreenUpdating = True
        
        # Let it recalculate
        print("Calculating...")
        start = time.perf_counter()
        excel.CalculateFull()
        print(f"  Calculated ({time.perf_counter() - start:.2f}s)")
        
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
