"""Restore Email headers validation array in VBA."""
import shutil, sys, tempfile
from pathlib import Path
import pythoncom, pywintypes, win32com.client as win32

def fix_workbook(path_str):
    path = Path(path_str).resolve()
    print(f"Fixing {path}...")
    temp_dir = Path(tempfile.mkdtemp())
    qa_path = temp_dir / path.name
    shutil.copy2(path, qa_path)
    
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    wb = excel.Workbooks.Open(str(qa_path), UpdateLinks=0)
    changed = False
    for i in range(1, wb.VBProject.VBComponents.Count + 1):
        comp = wb.VBProject.VBComponents(i)
        if comp.CodeModule.CountOfLines > 0:
            code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            orig = code
            
            # The email validation array is inside ValidateWorkbookConfiguration
            # We want to replace the first occurrence of:
            # "Segments", _
            # "Proof of Schedule", _
            # with the original Email array. But wait, we can just find where it's assigned to emailHeaders
            
            # The email headers are assigned:
            #    emailHeaders = Array( _
            #        ...
            #        "Approval", _
            #        "Segments", _
            #        "Proof of Schedule", _
            #        "Bluecore/Attentive Link", _
            #        "Est. Audience", _
            
            # We can search for the emailHeaders = Array block and replace inside it.
            if "emailHeaders = Array(" in code:
                idx1 = code.find("emailHeaders = Array(")
                idx2 = code.find("smsHeaders = Array(", idx1)
                
                email_part = code[idx1:idx2]
                fixed_email_part = email_part.replace(
                    '"Segments", _\r\n        "Proof of Schedule", _',
                    '"Segments", _\r\n        "Jira Link", _\r\n        "ClickUp Link", _'
                )
                
                code = code[:idx1] + fixed_email_part + code[idx2:]
            
            if code != orig:
                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                comp.CodeModule.AddFromString(code)
                changed = True
                print(f"  Fixed module {comp.Name}")
                
    if changed:
        wb.Save()
        print("  Saved changes.")
    wb.Close(SaveChanges=False)
    excel.Quit()
    shutil.copy2(qa_path, path)
    shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    fix_workbook(r"Production Tracker\Email & SMS Campaign Tracker.xlsm")
    fix_workbook(r"Production Tracker\Email & SMS Campaign Tracker Template.xlsm")
