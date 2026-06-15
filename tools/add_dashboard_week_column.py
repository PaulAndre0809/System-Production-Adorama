"""Add a "Week Number" column on the Dashboard, immediately right of the
DashboardWorkTable's "Bluecore/Attentive" column (column M), beneath the Week
Number KPI tile.

The workbook's embedded ValidateWorkbookConfiguration macro requires the
DashboardWorkTable to have exactly 12 columns, so the new column is added as a
standalone sheet column adjacent to the table (NOT a 13th table column). This
keeps every macro and the self-validation untouched. The table's own auto-expand
is disabled while writing so Excel does not absorb the column.

Each row shows a "Week N" label derived from the leading MMDDYY date code in the
Campaign cell (e.g. "061726-STO-Services-Trade-P-B-NA-GLP" -> "Week 25"), falling
back to the row's Send Date when the campaign has no parseable code (e.g. "TBD").
Blank feed rows stay blank. WEEKNUM(..,1) uses a Sunday-start week to match the
Dashboard feed and the Week Number tile.

All edits go through desktop Excel via win32com to preserve VBA, data validation,
tables and spill formulas. Idempotent: re-running removes any prior Week Number
table column (legacy) and overwrites the standalone column.

Usage:
    python tools/add_dashboard_week_column.py                 # all three workbooks
    python tools/add_dashboard_week_column.py <path> [<path>] # specific workbooks
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

TABLE = "DashboardWorkTable"
WEEK_COL = "Week Number"
COL = "M"            # immediately right of "Bluecore/Attentive" (column L)
HEADER_ROW = 10
FIRST_DATA = 11
LAST_DATA = 160      # matches DashboardWorkTable body A11:L160
XL_PASTE_FORMATS = -4122

# Relative A1 formula (anchored on row 11; Excel adjusts each row when filled down).
# Reads the row's Send Date (A) and Campaign (D), parses the leading MMDDYY code,
# falls back to Send Date, and emits "Week N".
WEEK_FORMULA = (
    "=LET("
    "sendDate,$A11,"
    'camp,$D11&"",'
    "code,LEFT(camp,6),"
    'parsed,IFERROR(DATE(2000+VALUE(MID(code,5,2)),VALUE(LEFT(code,2)),VALUE(MID(code,3,2))),""),'
    'useDate,IF(parsed="",IF(ISNUMBER(sendDate),INT(sendDate),""),parsed),'
    'IF(useDate="","","Week "&WEEKNUM(useDate,1)))'
)


def get_excel():
    try:
        return win32.gencache.EnsureDispatch("Excel.Application")
    except Exception:
        gen_path = getattr(win32com, "__gen_path__", None)
        if gen_path and os.path.isdir(gen_path):
            shutil.rmtree(gen_path, ignore_errors=True)
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


def add_week_column(ws):
    app = ws.Application
    was_protected = bool(retry_com(lambda: ws.ProtectContents))
    if was_protected:
        retry_com(lambda: ws.Unprotect())
    try:
        # idempotency / undo any legacy attempt: drop a Week Number *table* column
        lo = retry_com(lambda: ws.ListObjects(TABLE))
        for i in range(retry_com(lambda: lo.ListColumns.Count), 0, -1):
            if retry_com(lambda i=i: lo.ListColumns(i).Name) == WEEK_COL:
                retry_com(lambda i=i: lo.ListColumns(i).Delete())
        # keep Excel from absorbing the adjacent column into the table
        retry_com(lambda: setattr(app.AutoCorrect, "AutoExpandListRange", False))

        retry_com(lambda: setattr(ws.Range(f"{COL}{HEADER_ROW}"), "Value", WEEK_COL))
        body = f"{COL}{FIRST_DATA}:{COL}{LAST_DATA}"
        retry_com(lambda: setattr(ws.Range(body), "Formula2", WEEK_FORMULA))

        # match the look: header style from the Bluecore header, body banding from a
        # plain text column (Channel) so we don't inherit the hyperlink/date styling
        retry_com(lambda: ws.Range(f"L{HEADER_ROW}").Copy())
        retry_com(lambda: ws.Range(f"{COL}{HEADER_ROW}").PasteSpecial(Paste=XL_PASTE_FORMATS))
        retry_com(lambda: ws.Range(f"C{FIRST_DATA}:C{LAST_DATA}").Copy())
        retry_com(lambda: ws.Range(body).PasteSpecial(Paste=XL_PASTE_FORMATS))
        retry_com(lambda: setattr(app, "CutCopyMode", False))
        return retry_com(lambda: lo.Range.Address)
    finally:
        if was_protected:
            retry_com(lambda: ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True))


def process(excel, path: Path):
    print(f"  Opening {path.name}")
    wb = retry_com(lambda: excel.Workbooks.Open(str(path), UpdateLinks=0))
    try:
        name = wb.Name
        ref = add_week_column(wb.Worksheets("Dashboard"))
        retry_com(lambda: excel.CalculateFull())
        ws = wb.Worksheets("Dashboard")
        m10 = retry_com(lambda: ws.Range("M10").Value)
        m11 = retry_com(lambda: ws.Range("M11").Text)
        m14 = retry_com(lambda: ws.Range("M14").Text)
        validation = retry_com(lambda: excel.Run(f"'{name}'!ValidateWorkbookConfiguration"))
        print(f"    [Dashboard] table={ref} (unchanged); M10={m10!r}; M11={m11!r} M14={m14!r}; "
              f"ValidateWorkbookConfiguration={validation!r}")
        if validation != "OK":
            raise RuntimeError(f"embedded validation returned {validation!r}")
        retry_com(wb.Save)
        print(f"  Saved {path.name}")
    finally:
        retry_com(lambda: wb.Close(False))


def main():
    repo = Path(__file__).resolve().parents[1]
    folder = repo / "Production Tracker"
    if len(sys.argv) > 1:
        paths = [Path(a).resolve() for a in sys.argv[1:]]
    else:
        paths = [
            folder / "Email & SMS Campaign Tracker.xlsm",
            folder / "Email & SMS Campaign Tracker Template.xlsm",
            folder / "Email & SMS Campaign Tracker_backup.xlsm",
        ]
    for p in paths:
        if not p.exists():
            print(f"ERROR: missing workbook {p}")
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
        for name, err in failures:
            print(f" - {name}: {err}")
        return 1
    print("\nWEEK NUMBER COLUMN ADDED SUCCESSFULLY")
    return 0


if __name__ == "__main__":
    sys.exit(main())
