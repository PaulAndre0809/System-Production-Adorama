"""Fix the #VALUE! error in the campaign link columns caused by long URLs.

Root cause: the timed-link feature stores each URL as a string literal inside a
HYPERLINK formula. Excel stores string literals longer than 255 characters via the
internal `_xlfn._LONGTEXT(...)` wrapper, which evaluates to #VALUE! in this
workbook. Any link whose URL exceeds 255 characters therefore shows #VALUE!
(seen most often on the Bluecore/Attentive "compose/design?..." links).

Fix: patch the VBA so that when a link URL is longer than 255 characters, the cell
gets a REAL Excel hyperlink (whose address has no 255-char limit) showing the clean
platform name, instead of the formula. Short URLs keep the existing live timed
formula. Then run ApplyTimedCampaignLinks, which reprocesses every link column
(Email: Jira/ClickUp/Bluecore; SMS: Proof of Schedule/Bluecore) and repairs the
existing broken cells with the new logic. This also prevents recurrence, because the
same InstallTimedCampaignLink path runs whenever a user edits a link cell.

ValidateWorkbookConfiguration must return OK before saving. Idempotent.

Usage:
    python tools/fix_long_url_links.py [<path> ...]   # default: all three
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

VBA_OLD = (
    '    If LCase$(Left$(linkAddress, 4)) <> "http" Then Exit Sub\n'
    '\n'
    '    formulaText = TimedCampaignLinkFormula( _\n'
    '        linkAddress, displayName)\n'
)
VBA_NEW = (
    '    If LCase$(Left$(linkAddress, 4)) <> "http" Then Exit Sub\n'
    '\n'
    "    ' URLs longer than 255 characters cannot be stored as a formula string\n"
    "    ' literal (Excel wraps them in _xlfn._LONGTEXT, which evaluates to #VALUE!).\n"
    "    ' Use a real cell hyperlink instead so the link works for any URL length.\n"
    '    If Len(linkAddress) > 255 Then\n'
    '        InstallLongCampaignLink cell, linkAddress, displayName\n'
    '        Exit Sub\n'
    '    End If\n'
    '\n'
    '    formulaText = TimedCampaignLinkFormula( _\n'
    '        linkAddress, displayName)\n'
)

# appended after InstallTimedCampaignLink's End Sub
VBA_ANCHOR_END = (
    '    If CStr(cell.Formula) <> formulaText Then\n'
    '        ApplyFormulaCompat cell, formulaText\n'
    '    End If\n'
    'End Sub\n'
)
VBA_NEW_SUB = (
    '    If CStr(cell.Formula) <> formulaText Then\n'
    '        ApplyFormulaCompat cell, formulaText\n'
    '    End If\n'
    'End Sub\n'
    '\n'
    'Private Sub InstallLongCampaignLink( _\n'
    '    ByVal cell As Range, _\n'
    '    ByVal linkAddress As String, _\n'
    '    ByVal displayName As String)\n'
    '\n'
    '    Dim priorEvents As Boolean\n'
    '\n'
    '    priorEvents = Application.EnableEvents\n'
    '    Application.EnableEvents = False\n'
    '    On Error Resume Next\n'
    '    cell.Hyperlinks.Delete\n'
    '    cell.ClearContents\n'
    '    cell.Parent.Hyperlinks.Add Anchor:=cell, _\n'
    '        Address:=linkAddress, TextToDisplay:=displayName\n'
    '    On Error GoTo 0\n'
    '    Application.EnableEvents = priorEvents\n'
    'End Sub\n'
)


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
    if "InstallLongCampaignLink" in code:
        print("    VBA already patched")
        return
    norm = code.replace("\r\n", "\n").replace("\r", "\n")
    for needle in (VBA_OLD, VBA_ANCHOR_END):
        if norm.count(needle) != 1:
            raise RuntimeError(f"expected exactly 1 of fragment: {needle[:50]!r} (found {norm.count(needle)})")
    norm = norm.replace(VBA_OLD, VBA_NEW)
    norm = norm.replace(VBA_ANCHOR_END, VBA_NEW_SUB)
    retry_com(lambda: cm.DeleteLines(1, count))
    retry_com(lambda: cm.AddFromString(norm.replace("\n", "\r\n")))
    print("    VBA patched (long-URL links -> real hyperlink)")


def process(excel, path: Path):
    print(f"  Opening {path.name}")
    wb = retry_com(lambda: excel.Workbooks.Open(str(path), UpdateLinks=0))
    try:
        name = wb.Name
        edit_vba(wb)
        # reprocess every link column with the patched logic (repairs broken cells)
        retry_com(lambda: excel.Run(f"'{name}'!ApplyTimedCampaignLinks"))
        retry_com(lambda: excel.CalculateFull())
        validation = retry_com(lambda: excel.Run(f"'{name}'!ValidateWorkbookConfiguration"))
        print(f"    ApplyTimedCampaignLinks done; ValidateWorkbookConfiguration={validation!r}")
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
