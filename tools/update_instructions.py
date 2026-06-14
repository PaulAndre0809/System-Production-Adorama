import sys
from pathlib import Path
import win32com.client as win32
import time

def retry(func, attempts=10, delay=0.5):
    for i in range(attempts):
        try:
            return func()
        except Exception as e:
            if i == attempts - 1:
                raise
            time.sleep(delay)

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
            print(f"Updating instructions for {p.name}")
            wb = retry(lambda: excel.Workbooks.Open(str(p), UpdateLinks=0))
            
            try:
                ws = wb.Worksheets("Notes - Instructions")
                
                ws.Range("B35").Value = "7. Dynamic Dashboard Highlighting (NEW)"
                ws.Range("B35").Font.Bold = True
                
                ws.Range("C36").Value = "The Dashboard queue automatically highlights campaigns scheduled for tomorrow (Mon-Thu) or the weekend (Fri) with a soft blue background to help surface immediate deliverables."
                
                wb.Save()
                print(f"Updated instructions in {p.name}")
            except Exception as e:
                print(f"Failed to update instructions in {p.name}: {e}")
                
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
