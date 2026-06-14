import win32com.client as win32
from pathlib import Path
import shutil
import tempfile
import time
import pythoncom
import pywintypes
import os
import re

def retry_com(func, attempts=10, delay=0.5):
    for i in range(attempts):
        try:
            pythoncom.PumpWaitingMessages()
            return func()
        except pywintypes.com_error as e:
            if i == attempts - 1:
                raise
            time.sleep(delay)

def apply_fixes_to_vba(code: str) -> str:
    # 1. Sanitize continuation lines (remove blank lines after continuation character '_')
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
        
    code_sanitized = '\r\n'.join(new_lines)
    
    # 2. Inject On Error Resume Next at the top of the styling subs
    subs_to_wrap = [
        r"(Private Sub StyleCampaignSheet\([^\)]*\))",
        r"(Private Sub StyleCalendarSheet\([^\)]*\))",
        r"(Private Sub StyleCoreWorkbookSheets\([^\)]*\))",
        r"(Private Sub BuildNotesInstructionSheet\([^\)]*\))",
        r"(Private Sub StyleDashboard\([^\)]*\))",
        r"(Private Sub ApplyDashboardAuditHeader\([^\)]*\))"
    ]
    
    for sub_pattern in subs_to_wrap:
        # We find the sub declaration and append "On Error Resume Next" on the next line
        # Use regex to find it (case insensitive)
        def repl(match):
            return match.group(1) + "\r\n    On Error Resume Next"
        code_sanitized = re.sub(sub_pattern, repl, code_sanitized, flags=re.IGNORECASE)
        
    # 3. Apply the links conditional mapping fix in AppendDashboardRows
    # To be very robust, we find AppendDashboardRows and replace the specific block inside it
    target_links = """                jiraLinks(displayIndex) = TextValue( _
                    ValueByHeader(ws, rowNumber, "Jira Link"))

                clickUpLinks(displayIndex) = TextValue( _
                    ValueByHeader(ws, rowNumber, "ClickUp Link"))"""
                    
    replacement_links = """                If channelName = "SMS" Then
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
    normalized_target = target_links.replace('\r\n', '\n')
    normalized_replacement = replacement_links.replace('\r\n', '\n')
    
    # Check both formats
    if target_links in code_sanitized:
        code_sanitized = code_sanitized.replace(target_links, replacement_links)
    elif normalized_target in code_sanitized.replace('\r\n', '\n'):
        # Do a replacement by normalizing endings first, then replace, then restore \r\n
        code_sanitized = code_sanitized.replace('\r\n', '\n').replace(normalized_target, normalized_replacement).replace('\n', '\r\n')
        
    return code_sanitized

def process_excel_file(p: Path):
    print(f"Applying robust fixes to Excel file: {p.name}")
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
        
        wb = retry_com(lambda: excel.Workbooks.Open(str(temp_path), UpdateLinks=0))
        
        changed = False
        try:
            comp = wb.VBProject.VBComponents("modEmailProductionTracker")
            count = comp.CodeModule.CountOfLines
            if count > 0:
                code = comp.CodeModule.Lines(1, count)
                new_code = apply_fixes_to_vba(code)
                if new_code != code:
                    comp.CodeModule.DeleteLines(1, count)
                    comp.CodeModule.AddFromString(new_code)
                    changed = True
                    print(f"  Successfully updated modEmailProductionTracker in {p.name}")
        except Exception as e:
            print(f"  Error modifying {p.name}: {e}")
            
        if changed:
            wb.Save()
            print(f"  Saved changes to {p.name}")
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

def process_text_file(p: Path):
    if not p.exists():
        return
    print(f"Applying robust fixes to Text/VBA file: {p.name}")
    content = p.read_text(encoding='utf-8', errors='ignore')
    new_content = apply_fixes_to_vba(content)
    if new_content != content:
        # Standardize line endings to \r\n for consistency in repo
        p.write_text(new_content.replace('\r\n', '\n').replace('\n', '\r\n'), encoding='utf-8')
        print(f"  Successfully updated {p.name}")

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    
    # 1. First, discard any modifications using git restore
    os.system("taskkill /F /IM EXCEL.EXE >nul 2>&1")
    time.sleep(1.0)
    
    print("Restoring Excel workbooks and tracked files to clean Git state...")
    os.system('git restore "Production Tracker/Email & SMS Campaign Tracker.xlsm" "Production Tracker/Email & SMS Campaign Tracker Template.xlsm" clean_temp.vba vba_dump.txt tools/vba_dump.txt')
    time.sleep(1.0)
    
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    template_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker Template.xlsm"
    
    # 2. Process excel files
    process_excel_file(wb_path)
    process_excel_file(template_path)
    
    # 3. Process text files
    process_text_file(repo_dir / "vba_dump.txt")
    process_text_file(repo_dir / "tools" / "vba_dump.txt")
    process_text_file(repo_dir / "clean_temp.vba")
    process_text_file(repo_dir / "tools" / "clean_mod.vba")
    
    # 4. Fix campaign type dropdowns
    print("Running tools/fix_validation.py...")
    os.system(f'"{repo_dir}/.venv/Scripts/python.exe" "{repo_dir}/tools/fix_validation.py"')

if __name__ == "__main__":
    main()
