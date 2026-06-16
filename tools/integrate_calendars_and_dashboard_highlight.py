"""Integrate the SharePoint-linked monthly calendars and add Dashboard schedule-gap
highlighting.

For each workbook processed:
1. VBA (modEmailProductionTracker):
   - Remove the ValidateWorkbookConfiguration check that flags any "Calendar" sheet
     as a retired sheet (it now fails on the live SharePoint calendars).
   - Re-point the Notes generator's "Retired Features" row to a "Monthly Calendars"
     row, and drop the "Calendar" wording from the Maintenance note, so a Notes
     rebuild no longer reintroduces the retired-calendar guidance.
2. Dashboard: add two native conditional-formatting rules to the feed (A11:L160) that
   highlight schedule-gap rows - Email rows orange, SMS rows yellow - when the row's
   Send Date is in the next working window (tomorrow; Fridays extend through Mon) and
   its Stage is not yet "Scheduled" or "Sent" (cancelled rows excluded). Mirrors the
   source-sheet schedule-gap highlighting.
3. Notes - Instructions: replace the "Retired Features" entry (row 11) with a
   "Monthly Calendars" entry documenting the SharePoint mirror + duplication
   procedure, and update the Maintenance compatibility note (row 14).

The embedded ValidateWorkbookConfiguration must return OK before saving.

Usage:
    python tools/integrate_calendars_and_dashboard_highlight.py <path> [<path> ...]
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

ORANGE = 49407  # #FFC000
YELLOW = 65535  # #FFFF00
CF_RANGE = "A11:L160"
CF_SIG = "WEEKDAY(TODAY(),1)=6"


def cf_formula(channel: str) -> str:
    return (
        "=AND("
        f'$C11="{channel}",'
        "IFERROR(INT($A11),0)>=TODAY()+1,"
        "IFERROR(INT($A11),0)<=TODAY()+IF(WEEKDAY(TODAY(),1)=6,3,1),"
        '$F11<>"Scheduled",$F11<>"Sent",'
        'LEFT($D11,9)<>"CANCELLED")'
    )


# ---- VBA edits (exact live fragments, normalised to \n) ----
VBA_EDITS = [
    (  # 1: remove the retired-calendar validation loop
        "    ' Confirm that retired Calendar sheets cannot reappear unnoticed.\n"
        "    For Each candidateSheet In ThisWorkbook.Worksheets\n"
        '        If InStr(1, candidateSheet.Name, "Calendar", vbTextCompare) > 0 Then\n'
        "            Err.Raise Number:=vbObjectError + 1087, Description:= _\n"
        '                "Retired calendar sheet still exists: " & candidateSheet.Name\n'
        "        End If\n"
        "    Next candidateSheet\n",
        "    ' Monthly SharePoint-linked Calendar sheets are an active feature (no longer retired).\n",
    ),
    (  # 2: Notes generator row
        '    AddInstructionRow ws, nextRow, "Retired Features", "Calendars and weekly comparisons", _\n'
        '        "Monthly Calendar sheets and the Last Week vs Current Week Email/SMS comparison tables and charts were intentionally removed.", _\n'
        '        "Do not recreate sheets with Calendar in the name or Dashboard comparison objects named DeliveredComparison.", _\n'
        '        "Legacy macro names remain as safe compatibility wrappers and will not recreate retired features."',
        '    AddInstructionRow ws, nextRow, "Monthly Calendars", "View SharePoint planning calendars", _\n'
        '        "June and May 2026 Calendar sheets mirror the team SharePoint Email and SMS planning files through external-link formulas. Template for Duplicate is a pre-formatted month for creating new calendars.", _\n'
        '        "To add a month, copy Template for Duplicate, rename it as the month plus 2026 Calendar, then use Data, Edit Links to point its source to that month SharePoint files. Keep the calendar layout unchanged.", _\n'
        '        "Links refresh in desktop Excel with SharePoint access. Last Week vs Current Week comparison tables remain retired."',
    ),
    (  # 3: Maintenance note
        '        "The retired Calendar and comparison features remain removed after configuration refreshes."',
        '        "Last Week vs Current Week comparison tables remain retired after configuration refreshes."',
    ),
]

NOTE_ROW11 = [
    "Monthly Calendars",
    "View SharePoint planning calendars",
    ("June 2026 Calendar (visible) and May 2026 Calendar (hidden) mirror the team's SharePoint "
     "Email and SMS planning files through external-link formulas; no manual data entry is needed. "
     "Template for Duplicate is a pre-formatted month used to create new calendars."),
    ("To add a month: (1) right-click Template for Duplicate and Move or Copy to a new sheet; "
     "(2) rename it as the month plus 2026 Calendar; (3) Data > Edit Links > Change Source to that "
     "month's SharePoint Email and SMS planning files; (4) update. Keep the calendar layout and "
     "formatting unchanged."),
    ("Links refresh in desktop Excel when you have access to the SharePoint files; choose Update "
     "Values when prompted. Last Week vs Current Week comparison tables remain retired."),
]
NOTE_E14 = "Last Week vs Current Week comparison tables remain retired after configuration refreshes."


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
    if "Monthly SharePoint-linked Calendar sheets are an active feature" in code:
        print("    VBA already updated - skipping")
        return
    norm = code.replace("\r\n", "\n").replace("\r", "\n")
    for old, new in VBA_EDITS:
        n = norm.count(old)
        if n != 1:
            raise RuntimeError(f"VBA fragment count {n} (expected 1): {old[:60]!r}")
        norm = norm.replace(old, new)
    retry_com(lambda: cm.DeleteLines(1, count))
    retry_com(lambda: cm.AddFromString(norm.replace("\n", "\r\n")))
    print("    VBA edited (removed calendar check, updated Notes generator)")


def add_dashboard_cf(wb):
    ws = retry_com(lambda: wb.Worksheets("Dashboard"))
    was = bool(retry_com(lambda: ws.ProtectContents))
    if was:
        retry_com(lambda: ws.Unprotect())
    try:
        rng = retry_com(lambda: ws.Range(CF_RANGE))
        # idempotency: drop prior copies of our two rules
        for i in range(retry_com(lambda: rng.FormatConditions.Count), 0, -1):
            f = retry_com(lambda i=i: rng.FormatConditions(i).Formula1)
            if f and CF_SIG in f and ('"Email"' in f or '"SMS"' in f):
                retry_com(lambda i=i: rng.FormatConditions(i).Delete())
        for channel, color in (("Email", ORANGE), ("SMS", YELLOW)):
            fc = retry_com(lambda ch=channel: rng.FormatConditions.Add(Type=2, Formula1=cf_formula(ch)))
            retry_com(lambda fc=fc: setattr(fc, "StopIfTrue", False))
            retry_com(lambda fc=fc, color=color: setattr(fc.Interior, "Color", color))
            retry_com(lambda fc=fc: setattr(fc.Font, "Color", 0))
        print(f"    Dashboard schedule-gap CF added on {CF_RANGE} (Email=orange, SMS=yellow)")
    finally:
        if was:
            retry_com(lambda: ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True))


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
        # row 11 must currently be the Retired Features entry (or already updated)
        a11 = str(retry_com(lambda: ws.Range("A11").Value) or "")
        if a11 in ("Retired Features", "Monthly Calendars"):
            for c, val in enumerate(NOTE_ROW11, start=1):
                retry_com(lambda c=c, val=val: setattr(ws.Cells(11, c), "Value", val))
        else:
            print(f"    WARNING: Notes row 11 is {a11!r}; skipping row 11 rewrite")
        if str(retry_com(lambda: ws.Range("A14").Value) or "") == "Maintenance":
            retry_com(lambda: setattr(ws.Range("E14"), "Value", NOTE_E14))
        retry_com(lambda: ws.Range("A11:E14").EntireRow.AutoFit())
        print("    Notes - Instructions updated (row 11 -> Monthly Calendars; row 14 maintenance)")
    finally:
        retry_com(lambda: ws.Protect(Password=used_pw, DrawingObjects=True, Contents=True, Scenarios=True))


def process(excel, path: Path):
    print(f"  Opening {path.name}")
    wb = retry_com(lambda: excel.Workbooks.Open(str(path), UpdateLinks=0))
    try:
        name = wb.Name
        edit_vba(wb)
        add_dashboard_cf(wb)
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
        for name, err in failures:
            print(f" - {name}: {err}")
        return 1
    print("\nDONE")
    return 0


if __name__ == "__main__":
    sys.exit(main())
