"""Apply schedule-gap conditional formatting and a Dashboard Week Number tile.

Adds, to each of the three Email & SMS Campaign Tracker workbooks:

1. Native conditional formatting that highlights *unscheduled* deployments due in
   the next working window:
       - Email Campaigns rows -> ORANGE
       - SMS Campaigns rows   -> YELLOW
   "Unscheduled" = the row's `Scheduled` checkbox is not TRUE. "Next working
   window" = tomorrow, except on Fridays it extends through Sat/Sun/Mon. Cancelled
   rows are excluded. Rules are native (not VBA fills) so they recalc daily and
   work in Excel for the web.

2. A "Week Number" KPI-style tile on the Dashboard (M4:N6) showing the current
   two-week window as a span, e.g. "25-26".

3. A documentation section appended to the protected "Notes - Instructions" sheet.

All edits go through desktop Excel via win32com so VBA, data validation, in-cell
checkboxes, tables and spill formulas are preserved (openpyxl would strip the
data-validation extension). The script is idempotent: re-running replaces the
rules/tile/notes it previously added instead of duplicating them.
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

# --- Excel enum constants (literal so late-binding DispatchEx is not required) ---
XL_EXPRESSION = 2          # xlExpression
XL_PASTE_FORMATS = -4122   # xlPasteFormats
XL_UP = -4162              # xlUp

# --- Highlight fills (Excel BGR long = R + G*256 + B*65536) ---
ORANGE = 255 + 192 * 256 + 0 * 65536   # #FFC000  -> 49407
YELLOW = 255 + 255 * 256 + 0 * 65536   # #FFFF00  -> 65535
BLACK = 0

# Unique substring that identifies *our* conditional-format rule, used to find and
# remove a previous copy on re-runs without touching other rules.
CF_SIGNATURE = "WEEKDAY(TODAY(),1)=6"

NOTES_SHEET = "Notes - Instructions"
# The three workbooks were protected with different passwords (active/template use the
# first, the older backup uses the second). Try each and re-protect with whichever
# unlocked the sheet so no workbook's password is changed. Both verified against the
# stored SHA-512 protection hashes.
NOTES_PASSWORDS = ["Adorama@042026_", "adorama2024"]


def date_window_clause(scheduled_col: str, notes_col: str) -> str:
    """Build the shared schedule-gap test, anchored on table row 2.

    True when: the row has a Campaign Name, the Scheduled box is not TRUE, the
    Send Date falls in [tomorrow .. tomorrow (or +3 on Fridays)], and the row is
    not cancelled.
    """
    return (
        "=AND("
        "$C2<>\"\","
        f"${scheduled_col}2<>TRUE,"
        "IFERROR(INT($A2),0)>=TODAY()+1,"
        "IFERROR(INT($A2),0)<=TODAY()+IF(WEEKDAY(TODAY(),1)=6,3,1),"
        f"NOT(OR(LOWER(TRIM(${notes_col}2))=\"cancelled\",LOWER(TRIM(${notes_col}2))=\"canceled\"))"
        ")"
    )


def get_excel():
    """Early-bound Excel via makepy so optional/keyword COM args resolve correctly
    (dynamic DispatchEx raises DISP_E_PARAMNOTFOUND on FormatConditions.Add etc.)."""
    try:
        return win32.gencache.EnsureDispatch("Excel.Application")
    except Exception:
        gen_path = getattr(win32com, "__gen_path__", None)
        if gen_path and os.path.isdir(gen_path):
            shutil.rmtree(gen_path, ignore_errors=True)
        return win32.gencache.EnsureDispatch("Excel.Application")


def retry_com(func, attempts=12, delay=0.5):
    last = None
    for i in range(attempts):
        try:
            pythoncom.PumpWaitingMessages()
            return func()
        except pywintypes.com_error as exc:  # transient COM "server busy" etc.
            last = exc
            if i == attempts - 1:
                raise
            time.sleep(delay)
    raise last


def remove_existing_signature_rules(ws, cf_range: str) -> int:
    """Delete any prior copy of our rule on cf_range (idempotency). Other rules
    (cancelled, current-stage) are left untouched via the signature filter."""
    rng = ws.Range(cf_range)
    removed = 0
    # iterate from the end because deleting renumbers the collection
    for idx in range(rng.FormatConditions.Count, 0, -1):
        fc = rng.FormatConditions(idx)
        try:
            formula = fc.Formula1
        except Exception:
            formula = ""
        if formula and CF_SIGNATURE in formula:
            fc.Delete()
            removed += 1
    return removed


def apply_conditional_format(ws, cf_range: str, scheduled_col: str, notes_col: str, color: int):
    removed = remove_existing_signature_rules(ws, cf_range)
    rng = ws.Range(cf_range)
    formula = date_window_clause(scheduled_col, notes_col)
    fc = rng.FormatConditions.Add(Type=XL_EXPRESSION, Formula1=formula)
    fc.SetFirstPriority()          # whole-row highlight wins over per-column rules
    fc.StopIfTrue = False
    fc.Interior.Color = color
    fc.Font.Color = BLACK
    print(f"    [{ws.Name}] CF on {cf_range} (replaced {removed}) color={color:#08x}")


def apply_week_number_tile(wb):
    ws = wb.Worksheets("Dashboard")
    was_protected = bool(ws.ProtectContents)
    if was_protected:
        ws.Unprotect()  # Dashboard has no password
    try:
        # match the "Sent" tile styling (K4:L6) then lay out M4:N6
        ws.Range("K4:L6").Copy()
        ws.Range("M4:N6").PasteSpecial(Paste=XL_PASTE_FORMATS)
        ws.Application.CutCopyMode = False
        for row in (4, 5, 6):
            cell = ws.Range(f"M{row}:N{row}")
            if not cell.MergeCells:
                cell.Merge()
        ws.Range("M4").Value = "Week Number"
        ws.Range("M5").Formula = '=WEEKNUM(TODAY(),1)&"-"&WEEKNUM(TODAY()+7,1)'
        ws.Range("M6").Value = "Current + next week (Sun-Sat)"
        print(f"    [Dashboard] Week Number tile M4:N6 -> {ws.Range('M5').Text}")
    finally:
        if was_protected:
            # restore original protection (sheet=1 objects=1 scenarios=1, no password)
            ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True)


NOTE_ROWS = [
    [
        "Schedule-Gap Highlighting",
        "Spot unscheduled near-term sends",
        ("Email Campaigns rows turn orange and SMS Campaigns rows turn yellow when a "
         "campaign's Send Date is the next day and its Scheduled box is still unchecked. "
         "On Fridays the window extends through Saturday, Sunday, and Monday. Cancelled "
         "rows are never highlighted."),
        ("Check the Scheduled box once a deployment is set and the highlight clears "
         "automatically. Do not delete these conditional-formatting rules or rename the "
         "Scheduled, Send Date, Campaign Name, or Notes columns."),
        ("Native conditional formatting recalculates with the system date and works in "
         "desktop Excel and Excel for the web."),
    ],
    [
        "Dashboard Week Number",
        "Read the current two-week window",
        ("The Week Number tile beside the summary KPIs shows the current week through next "
         "week as a span, for example 25-26, using WEEKNUM(TODAY(),1) for a Sunday-start "
         "week that matches the Dashboard feed."),
        ("Display only. Do not overwrite the formula in the Week Number tile (Dashboard "
         "cell M5)."),
        ("Updates automatically with the system date in desktop Excel and Excel for the "
         "web."),
    ],
]


def unprotect_notes(ws):
    """Unlock the Notes sheet with whichever known password works; return it so the
    same password can be restored afterwards."""
    last = None
    for pw in NOTES_PASSWORDS:
        try:
            ws.Unprotect(Password=pw)
            return pw
        except pywintypes.com_error as exc:
            last = exc
    raise RuntimeError(f"Could not unprotect {NOTES_SHEET}; tried {NOTES_PASSWORDS}: {last!r}")


def apply_notes_section(wb):
    ws = wb.Worksheets(NOTES_SHEET)
    used_pw = unprotect_notes(ws)
    try:
        last_row = ws.Cells(ws.Rows.Count, 1).End(XL_UP).Row
        # idempotency: if our section is already present, overwrite in place
        existing = None
        for r in range(4, last_row + 1):
            if str(ws.Cells(r, 1).Value or "").strip() == NOTE_ROWS[0][0]:
                existing = r
                break
        start = existing if existing else last_row + 1
        template_row = 4  # copy the format of an existing entry row
        for offset, values in enumerate(NOTE_ROWS):
            r = start + offset
            ws.Range(f"A{template_row}:E{template_row}").Copy()
            ws.Range(f"A{r}:E{r}").PasteSpecial(Paste=XL_PASTE_FORMATS)
            ws.Application.CutCopyMode = False
            for c, val in enumerate(values, start=1):
                ws.Cells(r, c).Value = val
        ws.Range(f"A{start}:E{start + len(NOTE_ROWS) - 1}").EntireRow.AutoFit()
        print(f"    [{NOTES_SHEET}] wrote {len(NOTE_ROWS)} rows at {start}"
              f" ({'updated' if existing else 'appended'})")
    finally:
        # restore original protection options + the same password the sheet used
        ws.Protect(Password=used_pw, DrawingObjects=True, Contents=True, Scenarios=True)


def process_workbook(excel, path: Path):
    print(f"  Opening {path.name}")
    wb = retry_com(lambda: excel.Workbooks.Open(str(path), UpdateLinks=0))
    try:
        apply_conditional_format(wb.Worksheets("Email Campaigns"), "A2:W207", "O", "W", ORANGE)
        apply_conditional_format(wb.Worksheets("SMS Campaigns"), "A2:R201", "K", "R", YELLOW)
        apply_week_number_tile(wb)
        apply_notes_section(wb)
        retry_com(wb.Save)
        print(f"  Saved {path.name}")
    finally:
        wb.Close(SaveChanges=False)


def main():
    repo = Path(__file__).resolve().parents[1]
    folder = repo / "Production Tracker"
    if len(sys.argv) > 1:  # process only the workbooks passed on the command line
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

    # ensure no stale Excel instance holds a lock (repo tooling convention)
    os.system("taskkill /F /IM EXCEL.EXE >nul 2>&1")
    time.sleep(1.0)

    excel = None
    failures = []
    try:
        excel = get_excel()
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False          # keep Workbook_Open / Change handlers quiet
        excel.AutomationSecurity = 1        # msoAutomationSecurityLow (no macro prompt)
        for p in paths:
            try:
                process_workbook(excel, p)
            except Exception as exc:  # noqa: BLE001 - report and continue per-file
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
    print("\nALL THREE WORKBOOKS UPDATED SUCCESSFULLY")
    return 0


if __name__ == "__main__":
    sys.exit(main())
