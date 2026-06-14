import win32com.client as win32
from pathlib import Path
import time
import pythoncom
import pywintypes

def sanitize_vba_code(code):
    lines = code.splitlines()
    new_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        new_lines.append(line)
        if line.rstrip().endswith('_'):
            while i + 1 < len(lines) and not lines[i + 1].strip():
                i += 1
        i += 1
    return '\r\n'.join(new_lines)

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
            print(f"Sanitizing VBA in {p.name}")
            wb = retry(lambda: excel.Workbooks.Open(str(p), UpdateLinks=0))
            
            try:
                vb_project = wb.VBProject
                for j in range(1, vb_project.VBComponents.Count + 1):
                    comp = vb_project.VBComponents(j)
                    count = comp.CodeModule.CountOfLines
                    if count > 0:
                        code = comp.CodeModule.Lines(1, count)
                        sanitized = sanitize_vba_code(code)
                        if sanitized != code:
                            print(f"  Sanitizing module {comp.Name} ({count} lines)")
                            comp.CodeModule.DeleteLines(1, count)
                            if sanitized:
                                comp.CodeModule.AddFromString(sanitized)
                
                wb.Save()
                print(f"Successfully sanitized {p.name}")
            except Exception as e:
                print(f"Failed to sanitize {p.name}: {e}")
                
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
