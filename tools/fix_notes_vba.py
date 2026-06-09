"""Fix Notes - Instructions sheet name in VBA."""
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
            
            code = code.replace(
                'Public Const SH_INSTRUCTIONS As String = "Notes - Instruction"',
                'Public Const SH_INSTRUCTIONS As String = "Notes - Instructions"'
            )
            
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
