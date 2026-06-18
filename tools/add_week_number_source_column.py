"""Add a "Week Number" column to the left of "Send Date" in the Email Campaigns and
SMS Campaigns tables.

It is inserted as the first table column (so it sits immediately left of Send Date).
Each row shows a "Week N" label worked out from that row's Send Date using
WEEKNUM(Send Date, 1) - a Sunday-start week, matching the rest of the workbook - so
campaigns can be labelled and filtered by the week they belong to. Empty rows stay
blank.

Inserting the column shifts the other columns right; Excel automatically re-points the
existing conditional formatting (schedule-gap, Current Stage, Cancelled) and the
Dashboard's structured-reference formulas, so nothing else needs editing. The embedded
ValidateWorkbookConfiguration must return OK before saving.

Usage:
    python tools/add_week_number_source_column.py [<path> ...]   # default: all three
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

WEEK_COL = "Week Number"
FORMULA = '=IF(ISNUMBER([@[Send Date]]),"Week "&WEEKNUM([@[Send Date]],1),"")'
TABLES = [("Email Campaigns", "EmailCampaignsTable"),
          ("SMS Campaigns", "SMSCampaignsTable")]


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


def add_week_column(wb):
    for sheet, tbl in TABLES:
        ws = retry_com(lambda s=sheet: wb.Worksheets(s))
        lo = retry_com(lambda t=tbl: ws.ListObjects(t))
        first = retry_com(lambda: lo.ListColumns(1).Name)
        if first == WEEK_COL:                       # idempotent: refresh formula only
            col = retry_com(lambda: lo.ListColumns(1))
            retry_com(lambda: setattr(col.DataBodyRange, "Formula2", FORMULA))
            print(f"    [{sheet}] Week Number column already first - formula refreshed")
            continue
        col = retry_com(lambda: lo.ListColumns.Add(1))   # insert as the first column
        retry_com(lambda: setattr(col, "Name", WEEK_COL))
        retry_com(lambda: setattr(col.DataBodyRange, "Formula2", FORMULA))
        retry_com(lambda: setattr(col.Range, "ColumnWidth", 12))
        print(f"    [{sheet}] inserted Week Number as column A -> table {retry_com(lambda: lo.Range.Address)}")


def process(excel, path: Path):
    print(f"  Opening {path.name}")
    wb = retry_com(lambda: excel.Workbooks.Open(str(path), UpdateLinks=0))
    try:
        name = wb.Name
        add_week_column(wb)
        retry_com(lambda: excel.CalculateFull())
        e = retry_com(lambda: wb.Worksheets("Email Campaigns"))
        sample = retry_com(lambda: e.Cells(2, 1).Value)
        validation = retry_com(lambda: excel.Run(f"'{name}'!ValidateWorkbookConfiguration"))
        print(f"    sample Email A2={sample!r}; ValidateWorkbookConfiguration={validation!r}")
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
