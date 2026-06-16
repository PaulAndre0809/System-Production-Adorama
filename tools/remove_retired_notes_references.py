"""Remove the remaining "retired feature" references from the Notes - Instructions
sheet (and the VBA generator that produces it) so the documentation only describes
features that are still in use.

The only retired-feature wording left was the "Last Week vs Current Week comparison
tables remain retired" sentence in two places:
- Notes row 11 (Monthly Calendars) compatibility note  -> cell E11
- Notes row 14 (Maintenance) compatibility note         -> cell E14
plus the same strings inside the AddInstructionRow calls in modEmailProductionTracker.

Both are rewritten to active, non-retired wording. ValidateWorkbookConfiguration
must return OK before saving.

Usage:
    python tools/remove_retired_notes_references.py [<path> ...]   # default: all three
"""

import os
import shutil
import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com
import win32com.client as win32

MODULE = "modEmailProductionTracker"
NOTES_PASSWORDS = ["Adorama@042026_", "adorama2024"]

NEW_E11 = ("Links refresh in desktop Excel when you have access to the SharePoint files; "
           "choose Update Values when prompted.")
NEW_E14 = ("Configuration refreshes repair Dashboard formulas and formatting without "
           "rebuilding source tables or the SharePoint calendars.")

VBA_EDITS = [
    ('        "Links refresh in desktop Excel with SharePoint access. Last Week vs Current Week comparison tables remain retired."',
     '        "Links refresh in desktop Excel with SharePoint access; choose Update Values when prompted."'),
    ('        "Last Week vs Current Week comparison tables remain retired after configuration refreshes."',
     '        "Configuration refreshes repair Dashboard formulas and formatting without rebuilding source tables or the SharePoint calendars."'),
]


def get_excel():
    try:
        return win32.gencache.EnsureDispatch("Excel.Application")
    except Exception:
        gp = getattr(win32com, "__gen_path__", None)
        if gp and os.path.isdir(gp):
            shutil.rmtree(gp, ignore_errors=True)
        return win32.gencache.EnsureDispatch("Excel.Application")


def retry_com(func, attempts=20, delay=0.5):
    last = None
    for i in range(attempts):
        try:
            pythoncom.PumpWaitingMessages()
            return func()
        except pywintypes.com_error as exc:
            last = exc
            if i == attempts - 1:
                raise
            time.sleep(delay)
    raise last


def edit_vba(wb):
    cm = retry_com(lambda: wb.VBProject.VBComponents(MODULE).CodeModule)
    count = retry_com(lambda: cm.CountOfLines)
    code = retry_com(lambda: cm.Lines(1, count))
    norm = code.replace("\r\n", "\n").replace("\r", "\n")
    changed = False
    for old, new in VBA_EDITS:
        if old in norm:
            norm = norm.replace(old, new)
            changed = True
    if changed:
        retry_com(lambda: cm.DeleteLines(1, count))
        retry_com(lambda: cm.AddFromString(norm.replace("\n", "\r\n")))
        print("    VBA Notes generator cleaned of retired-comparison wording")
    else:
        print("    VBA already clean")


def update_notes(wb):
    ws = retry_com(lambda: wb.Worksheets("Notes - Instructions"))
    used_pw = None
    for pw in NOTES_PASSWORDS:
        try:
            retry_com(lambda pw=pw: ws.Unprotect(Password=pw))
            used_pw = pw
            break
        except pywintypes.com_error:
            continue
    if used_pw is None:
        raise RuntimeError("could not unprotect Notes - Instructions")
    try:
        retry_com(lambda: setattr(ws.Range("E11"), "Value", NEW_E11))
        retry_com(lambda: setattr(ws.Range("E14"), "Value", NEW_E14))
        retry_com(lambda: ws.Range("E11:E14").EntireRow.AutoFit())
        print("    Notes E11/E14 cleaned of retired-comparison wording")
    finally:
        retry_com(lambda: ws.Protect(Password=used_pw, DrawingObjects=True, Contents=True, Scenarios=True))


def process(excel, path: Path):
    print(f"  Opening {path.name}")
    wb = retry_com(lambda: excel.Workbooks.Open(str(path), UpdateLinks=0))
    try:
        name = wb.Name
        edit_vba(wb)
        update_notes(wb)
        retry_com(lambda: excel.CalculateFull())
        validation = retry_com(lambda: excel.Run(f"'{name}'!ValidateWorkbookConfiguration"))
        print(f"    ValidateWorkbookConfiguration={validation!r}")
        if validation != "OK":
            raise RuntimeError(f"embedded validation returned {validation!r}")
        retry_com(wb.Save)
        print(f"  Saved {path.name}")
    finally:
        retry_com(lambda: wb.Close(False))


def main():
    repo = Path(__file__).resolve().parents[1]
    folder = repo / "Production Tracker"
    paths = ([Path(a).resolve() for a in sys.argv[1:]] if len(sys.argv) > 1
             else [folder / "Email & SMS Campaign Tracker.xlsm",
                   folder / "Email & SMS Campaign Tracker Template.xlsm",
                   folder / "Email & SMS Campaign Tracker_backup.xlsm"])
    for p in paths:
        if not p.exists():
            print(f"ERROR: missing {p}")
            return 1
    os.system("taskkill /F /IM EXCEL.EXE >nul 2>&1")
    time.sleep(1.0)
    excel = None
    failures = []
    try:
        excel = get_excel()
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1
        for p in paths:
            try:
                process(excel, p)
            except Exception as exc:  # noqa: BLE001
                failures.append((p.name, repr(exc)))
                print(f"  FAILED {p.name}: {exc!r}")
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
    if failures:
        print("\nCOMPLETED WITH ERRORS:")
        for n, e in failures:
            print(f" - {n}: {e}")
        return 1
    print("\nDONE")
    return 0


if __name__ == "__main__":
    sys.exit(main())
