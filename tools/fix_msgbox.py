import sys
import win32com.client as win32
from pathlib import Path

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
            wb = excel.Workbooks.Open(str(p), UpdateLinks=0)
            
            try:
                comp = wb.VBProject.VBComponents("modDynamicHighlighting")
                count = comp.CodeModule.CountOfLines
                if count > 0:
                    code = comp.CodeModule.Lines(1, count)
                    
                    code = code.replace('MsgBox "Error: \'Send Date\' column not found', 'Err.Raise vbObjectError + 9999, , "Error: Send Date column not found')
                    code = code.replace('MsgBox "An error occurred while applying highlights: " & Err.Description, vbCritical', 'Err.Raise vbObjectError + 9998, , "An error occurred while applying highlights: " & Err.Description')
                    code = code.replace('MsgBox "Error: \'DashboardWorkTable\' not found', 'Err.Raise vbObjectError + 9997, , "Error: DashboardWorkTable not found')
                    
                    comp.CodeModule.DeleteLines(1, count)
                    comp.CodeModule.AddFromString(code)
                    
                    wb.Save()
                    print(f"Saved modified VBA in {p.name}")
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
