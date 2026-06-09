"""Quick check on Proof of Schedule column type in SMS table."""
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
temp_dir = Path(tempfile.mkdtemp(prefix="check_"))
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

    sms_ws = workbook.Worksheets("SMS Campaigns")
    sms_table = sms_ws.ListObjects("SMSCampaignsTable")

    # Check Proof of Schedule column
    for col in sms_table.ListColumns:
        if "proof" in col.Name.lower():
            print(f"Column: {col.Name!r} (index={col.Index})")
            # Check first few values
            for r in range(1, min(6, col.DataBodyRange.Rows.Count + 1)):
                val = col.DataBodyRange.Cells(r, 1).Value
                fmt = col.DataBodyRange.Cells(r, 1).NumberFormat
                print(f"  Row {r}: value={val!r} type={type(val).__name__} format={fmt!r}")
            break

    # Also check if there's a VBA module that references calendar sheets
    print("\n--- VBA Module Names ---")
    for i in range(1, workbook.VBProject.VBComponents.Count + 1):
        comp = workbook.VBProject.VBComponents(i)
        code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines) if comp.CodeModule.CountOfLines > 0 else ""
        has_cal_ref = "Calendar" in code
        print(f"  {comp.Name} (type={comp.Type}, lines={comp.CodeModule.CountOfLines}, has calendar ref={has_cal_ref})")
        if has_cal_ref:
            # Print lines with Calendar
            for line_num in range(1, comp.CodeModule.CountOfLines + 1):
                line = comp.CodeModule.Lines(line_num, 1)
                if "Calendar" in line:
                    print(f"    L{line_num}: {line.strip()}")

    workbook.Close(SaveChanges=False)
finally:
    if excel:
        excel.Quit()
    shutil.rmtree(temp_dir, ignore_errors=True)
