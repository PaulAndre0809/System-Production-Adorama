"""Copy the example calendar sheets from the active tracker into the Template and
Backup workbooks.

Per the agreed sync scope, the Template/Backup get a practical example only:
the "Template for Duplicate" sheet plus one real calendar ("June 2026 Calendar").
Copying the sheets brings their SharePoint external-link definitions along, which
serve as a ready-made guide/placeholder for future months. (The hidden
"May 2026 Calendar" is intentionally not copied.)

Run integrate_calendars_and_dashboard_highlight.py on the targets afterwards to
apply the VBA/Dashboard/Notes changes and confirm validation.

Usage:
    python tools/sync_calendars_to_template_backup.py <source> <target> [<target> ...]
"""

import os
import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com
import win32com.client as win32

SHEETS = ["June 2026 Calendar", "Template for Duplicate"]


def get_excel():
    try:
        return win32.gencache.EnsureDispatch("Excel.Application")
    except Exception:
        gp = getattr(win32com, "__gen_path__", None)
        if gp and os.path.isdir(gp):
            import shutil
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


def sheet_names(wb):
    return [retry_com(lambda i=i: wb.Worksheets(i).Name)
            for i in range(1, retry_com(lambda: wb.Worksheets.Count) + 1)]


def copy_sheets(src, tgt):
    anchor = "Dropdowns" if "Dropdowns" in sheet_names(tgt) else sheet_names(tgt)[-1]
    for name in SHEETS:
        # idempotency: remove an existing copy in the target first
        for i in range(retry_com(lambda: tgt.Worksheets.Count), 0, -1):
            if retry_com(lambda i=i: tgt.Worksheets(i).Name) == name:
                retry_com(lambda i=i: tgt.Worksheets(i).Delete())
        retry_com(lambda n=name: src.Worksheets(n).Copy(Before=tgt.Worksheets(anchor)))
    print(f"    copied {SHEETS} -> now: {sheet_names(tgt)}")


def main():
    if len(sys.argv) < 3:
        print("usage: sync_calendars_to_template_backup.py <source> <target> [<target> ...]")
        return 1
    source = Path(sys.argv[1]).resolve()
    targets = [Path(a).resolve() for a in sys.argv[2:]]
    for p in [source, *targets]:
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
        src = retry_com(lambda: excel.Workbooks.Open(str(source), UpdateLinks=0))
        for tgt_path in targets:
            try:
                print(f"  Target {tgt_path.name}")
                tgt = retry_com(lambda: excel.Workbooks.Open(str(tgt_path), UpdateLinks=0))
                copy_sheets(src, tgt)
                retry_com(tgt.Save)
                retry_com(lambda: tgt.Close(False))
                print(f"  Saved {tgt_path.name}")
            except Exception as exc:  # noqa: BLE001
                failures.append((tgt_path.name, repr(exc)))
                print(f"  FAILED {tgt_path.name}: {exc!r}")
        retry_com(lambda: src.Close(False))
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
    print("\nSHEETS SYNCED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
