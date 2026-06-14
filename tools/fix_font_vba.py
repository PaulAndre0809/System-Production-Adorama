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

def replace_font_name_in_code(code: str) -> str:
    replacements = {
        'ws.Columns("A:R").Font.Name = "Aptos"': 'On Error Resume Next: ws.Columns("A:R").Font.Name = "Aptos": On Error GoTo 0',
        'ws.Range("A1:G12").Font.Name = "Aptos"': 'On Error Resume Next: ws.Range("A1:G12").Font.Name = "Aptos": On Error GoTo 0',
        'ws.Range("A1:E200").Font.Name = "Aptos"': 'On Error Resume Next: ws.Range("A1:E200").Font.Name = "Aptos": On Error GoTo 0',
        'ws.Range("A1:D20").Font.Name = "Aptos"': 'On Error Resume Next: ws.Range("A1:D20").Font.Name = "Aptos": On Error GoTo 0',
        'ws.UsedRange.Font.Name = "Aptos"': 'On Error Resume Next: ws.UsedRange.Font.Name = "Aptos": On Error GoTo 0',
        'lo.Range.Font.Name = "Aptos"': 'On Error Resume Next: lo.Range.Font.Name = "Aptos": On Error GoTo 0',
        
        # Strikethrough safe wrappers
        'ws.Range("A1:G12").Font.Strikethrough = False': 'On Error Resume Next: ws.Range("A1:G12").Font.Strikethrough = False: On Error GoTo 0',
        'lo.Range.Font.Strikethrough = False': 'On Error Resume Next: lo.Range.Font.Strikethrough = False: On Error GoTo 0'
    }
    
    new_code = code
    for target, replacement in replacements.items():
        new_code = new_code.replace(target, replacement)
    return new_code

def fix_excel_file(p: Path):
    print(f"Modifying Excel file: {p.name}")
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
                new_code = replace_font_name_in_code(code)
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
    print(f"Modifying Text/VBA file: {p.name}")
    content = p.read_text(encoding='utf-8', errors='ignore')
    new_content = replace_font_name_in_code(content)
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
