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
            print(f"Opening workbook: {p.name}")
            wb = retry(lambda: excel.Workbooks.Open(str(p), UpdateLinks=0))
            
            try:
                comp = wb.VBProject.VBComponents("modEmailProductionTracker")
                count = comp.CodeModule.CountOfLines
                if count > 0:
                    code = comp.CodeModule.Lines(1, count)
                    target = 'MsgBox "Daily digest created on Dashboard, starting at T2.", vbInformation'
                    replacement = "    If False Then\r\n        " + target + "\r\n    End If"
                    
                    if target in code:
                        code = code.replace(target, replacement)
                        comp.CodeModule.DeleteLines(1, count)
                        comp.CodeModule.AddFromString(code)
                        wb.Save()
                        print(f"Successfully wrapped MsgBox in {p.name}")
                    else:
                        print(f"Target MsgBox not found in {p.name}")
            except Exception as e:
                print(f"Failed to process {p.name}: {e}")
                
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
