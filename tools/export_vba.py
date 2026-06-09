"""Export VBA code from workbook to inspect calendar references."""
import shutil, tempfile, time, sys
from pathlib import Path
import pythoncom, pywintypes, win32com.client as win32

def retry(label, func, attempts=20, delay=0.75):
    last_error = None
    for _ in range(attempts):
        try:
            pythoncom.PumpWaitingMessages()
            return func()
        except pywintypes.com_error as exc:
            last_error = exc
            if exc.args and exc.args[0] in (-2147418111, -2147417846):
                time.sleep(delay)
                continue
            raise
    raise RuntimeError(f"{label} failed repeatedly: {last_error}")

path = Path(r"Production Tracker\Email & SMS Campaign Tracker.xlsm").resolve()
temp_dir = Path(tempfile.mkdtemp(prefix="vba_"))
qa_path = temp_dir / path.name
shutil.copy2(path, qa_path)

excel = None
try:
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass

    workbook = retry("open", lambda: excel.Workbooks.Open(str(qa_path), UpdateLinks=0, ReadOnly=True, AddToMru=False))

    for i in range(1, workbook.VBProject.VBComponents.Count + 1):
        comp = retry(f"comp {i}", lambda i=i: workbook.VBProject.VBComponents(i))
        name = retry(f"name {i}", lambda: comp.Name)
        count = retry(f"count {i}", lambda: comp.CodeModule.CountOfLines)
        comp_type = retry(f"type {i}", lambda: comp.Type)
        print(f"\n{'='*60}")
        print(f"MODULE: {name} (type={comp_type}, lines={count})")
        print(f"{'='*60}")
        if count > 0:
            code = retry(f"code {name}", lambda: comp.CodeModule.Lines(1, count))
            print(code)

    workbook.Close(SaveChanges=False)
finally:
    if excel:
        excel.Quit()
    shutil.rmtree(temp_dir, ignore_errors=True)
