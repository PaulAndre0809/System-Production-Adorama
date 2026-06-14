import win32com.client as win32
from pathlib import Path
import shutil
import tempfile
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

def replace_links_in_code(code: str) -> str:
    target = """                jiraLinks(displayIndex) = TextValue( _
                    ValueByHeader(ws, rowNumber, "Jira Link"))

                clickUpLinks(displayIndex) = TextValue( _
                    ValueByHeader(ws, rowNumber, "ClickUp Link"))"""
                    
    replacement = """                If channelName = "SMS" Then
                    jiraLinks(displayIndex) = ""
                    clickUpLinks(displayIndex) = TextValue( _
                        ValueByHeader(ws, rowNumber, "Proof of Schedule"))
                Else
                    jiraLinks(displayIndex) = TextValue( _
                        ValueByHeader(ws, rowNumber, "Jira Link"))
                    clickUpLinks(displayIndex) = TextValue( _
                        ValueByHeader(ws, rowNumber, "ClickUp Link"))
                End If"""
                
    # Normalize double newlines or single newlines
    normalized_target = target.replace('\r\n', '\n')
    normalized_replacement = replacement.replace('\r\n', '\n')
    
    # We will do replacement for both formats (with \r\n and with \n)
    new_code = code
    if target in new_code:
        new_code = new_code.replace(target, replacement)
    elif normalized_target in new_code.replace('\r\n', '\n'):
        # Do a replacement by normalizing endings first
        new_code = new_code.replace('\r\n', '\n').replace(normalized_target, normalized_replacement)
        
    return new_code

def fix_excel_file(p: Path):
    print(f"Modifying links in Excel file: {p.name}")
    temp_dir = Path(tempfile.mkdtemp())
    temp_path = temp_dir / p.name
    shutil.copy2(p, temp_path)
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1
        
        wb = retry(lambda: excel.Workbooks.Open(str(temp_path), UpdateLinks=0))
        
        changed = False
        try:
            comp = wb.VBProject.VBComponents("modEmailProductionTracker")
            count = comp.CodeModule.CountOfLines
            if count > 0:
                code = comp.CodeModule.Lines(1, count)
                new_code = replace_links_in_code(code)
                if new_code != code:
                    comp.CodeModule.DeleteLines(1, count)
                    comp.CodeModule.AddFromString(new_code)
                    changed = True
                    print(f"  Updated modEmailProductionTracker in {p.name}")
        except Exception as e:
            print(f"  Error modifying {p.name}: {e}")
            
        if changed:
            wb.Save()
            print(f"  Saved {p.name}")
        wb.Close(SaveChanges=False)
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass
        if temp_path.exists():
            shutil.copy2(temp_path, p)
        shutil.rmtree(temp_dir, ignore_errors=True)

def fix_text_file(p: Path):
    if not p.exists():
        return
    print(f"Modifying links in Text/VBA file: {p.name}")
    content = p.read_text(encoding='utf-8', errors='ignore')
    new_content = replace_links_in_code(content)
    if new_content != content:
        p.write_text(new_content, encoding='utf-8')
        print(f"  Updated {p.name}")

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    template_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker Template.xlsm"
    
    # Terminate Excel first to ensure file is not locked
    import os
    os.system("taskkill /F /IM EXCEL.EXE >nul 2>&1")
    time.sleep(1.0)
    
    fix_excel_file(wb_path)
    fix_excel_file(template_path)
    
    # Fix VBA text files
    fix_text_file(repo_dir / "vba_dump.txt")
    fix_text_file(repo_dir / "tools" / "vba_dump.txt")
    fix_text_file(repo_dir / "clean_temp.vba")
    fix_text_file(repo_dir / "tools" / "clean_mod.vba")

if __name__ == "__main__":
    main()
