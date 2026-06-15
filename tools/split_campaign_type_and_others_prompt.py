"""Campaign Type changes for the Email & SMS Campaign Tracker:

1. Split the "Loyalty & PLCC" dropdown option into two independent options,
   "Loyalty" and "PLCC". The list source grows from Dropdowns!A2:A9 to A2:A10.
2. Make selecting "Others" pop up a text input (VBA InputBox) so the user can
   type a custom campaign type; the value fills that row's cell only (it is not
   added to the dropdown list), matching the existing custom-value design.

Both are driven through desktop Excel via win32com so VBA, validation, tables and
spill formulas are preserved. The embedded VBA (modEmailProductionTracker) is the
source of truth for the option list and self-validation, so it is edited in place:
- CampaignTypeOptions(): split the option.
- CampaignTypeDropdownIsConfigured(): expect the new $A$2:$A$10 source range.
- HandleCampaignChange(): call PromptCustomCampaignType for the changed cells.
- PromptCustomCampaignType(): the new "Others" InputBox handler (event path only,
  so automation/QA with events disabled never shows the modal).
The Dropdowns sheet and the column validation are then rewritten to match, and the
embedded ValidateWorkbookConfiguration must return OK before saving.

Usage:
    python tools/split_campaign_type_and_others_prompt.py                 # all three
    python tools/split_campaign_type_and_others_prompt.py <path> [<path>] # specific
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

TYPES = ["Promo", "Services", "Loyalty", "PLCC", "Newsletters", "Events", "NPA", "Others"]
VALIDATION_FORMULA = "=Dropdowns!$A$2:$A$10"
INPUT_MESSAGE = (
    "Choose a type. To use a type that is not listed, choose Others and enter the "
    "custom value when prompted. Custom values apply to that row only and are not "
    "added to the list."
)
ERROR_MESSAGE = (
    "Custom values are allowed for this row, but they are not saved as dropdown "
    "options."
)

# ---- exact live-code fragments (normalised to \n) ----
OLD_OPTION = '        "Loyalty & PLCC", _'
NEW_OPTION = '        "Loyalty", _\n        "PLCC", _'

OLD_RANGE = '"$A$2:$A$9"'
NEW_RANGE = '"$A$2:$A$10"'

CHANGE_ANCHOR = (
    "    Set changedCells = Intersect(Target, lo.DataBodyRange)\n"
    "    If changedCells Is Nothing Then Exit Sub\n"
)
CHANGE_INJECT = CHANGE_ANCHOR + (
    "\n    ' Prompt for a custom Campaign Type when \"Others\" is selected.\n"
    "    PromptCustomCampaignType lo, changedCells\n"
)

OLD_INPUTMSG = (
    "        .InputMessage = _\n"
    "            \"Choose a type. If the type is not listed, choose Others, \" & _\n"
    "            \"then type the custom value in this cell. Custom values are not \" & _\n"
    "            \"added to the dropdown list.\""
)
NEW_INPUTMSG = (
    "        .InputMessage = _\n"
    "            \"Choose a type. To use a type that is not listed, choose Others \" & _\n"
    "            \"and enter the custom value when prompted. Custom values apply to \" & _\n"
    "            \"that row only and are not added to the list.\""
)

PROMPT_SUB = '''
Private Sub PromptCustomCampaignType( _
    ByVal lo As ListObject, _
    ByVal changedCells As Range)

    Dim typeColumnIndex As Long
    Dim cell As Range
    Dim customValue As String
    Dim priorEvents As Boolean

    On Error Resume Next
    typeColumnIndex = lo.ListColumns("Campaign Type").Range.Column
    On Error GoTo 0
    If typeColumnIndex = 0 Then Exit Sub

    For Each cell In changedCells.Cells
        If cell.Column = typeColumnIndex Then
            If StrComp(Trim$(CStr(cell.Value)), "Others", vbTextCompare) = 0 Then
                customValue = Trim$(InputBox( _
                    "Enter a custom campaign type for this campaign." & vbCrLf & _
                    "(Leave blank to keep ""Others"".)", _
                    "Custom Campaign Type"))
                If Len(customValue) > 0 Then
                    priorEvents = Application.EnableEvents
                    Application.EnableEvents = False
                    cell.Value = customValue
                    Application.EnableEvents = priorEvents
                End If
            End If
        End If
    Next cell
End Sub
'''


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


def edit_module_code(code: str) -> str:
    norm = code.replace("\r\n", "\n").replace("\r", "\n")
    checks = {
        OLD_OPTION: 1,
        OLD_RANGE: 1,
        CHANGE_ANCHOR: 1,
        OLD_INPUTMSG: 1,
    }
    for needle, want in checks.items():
        got = norm.count(needle)
        if got != want:
            raise RuntimeError(f"expected {want} occurrence(s) of {needle!r}, found {got}")
    norm = norm.replace(OLD_OPTION, NEW_OPTION)
    norm = norm.replace(OLD_RANGE, NEW_RANGE)
    norm = norm.replace(CHANGE_ANCHOR, CHANGE_INJECT)
    norm = norm.replace(OLD_INPUTMSG, NEW_INPUTMSG)
    norm = norm.rstrip("\n") + "\n\n" + PROMPT_SUB.strip("\n") + "\n"
    return norm.replace("\n", "\r\n")


def edit_vba(wb):
    comp = retry_com(lambda: wb.VBProject.VBComponents(MODULE))
    cm = comp.CodeModule
    count = retry_com(lambda: cm.CountOfLines)
    code = retry_com(lambda: cm.Lines(1, count))
    if "PromptCustomCampaignType" in code:
        print("    VBA already updated - skipping code edits")
        return
    newcode = edit_module_code(code)
    retry_com(lambda: cm.DeleteLines(1, count))
    retry_com(lambda: cm.AddFromString(newcode))
    print("    VBA edited (split option, $A$2:$A$10, Others prompt)")


def apply_dropdown_and_validation(wb):
    dd = retry_com(lambda: wb.Worksheets("Dropdowns"))
    retry_com(lambda: setattr(dd.Range("A1"), "Value", "Campaign Type"))
    retry_com(lambda: dd.Range("A2:A30").ClearContents())
    for i, t in enumerate(TYPES):
        retry_com(lambda i=i, t=t: setattr(dd.Cells(2 + i, 1), "Value", t))  # A2..A9; A10 blank
    for sheet, tbl in [("Email Campaigns", "EmailCampaignsTable"),
                       ("SMS Campaigns", "SMSCampaignsTable")]:
        ws = retry_com(lambda s=sheet: wb.Worksheets(s))
        col = retry_com(lambda t=tbl: ws.ListObjects(t).ListColumns("Campaign Type"))
        v = retry_com(lambda: col.DataBodyRange.Validation)
        retry_com(lambda: v.Delete())
        retry_com(lambda: v.Add(Type=3, AlertStyle=3, Operator=1, Formula1=VALIDATION_FORMULA))
        retry_com(lambda: setattr(v, "IgnoreBlank", True))
        retry_com(lambda: setattr(v, "InCellDropdown", True))
        retry_com(lambda: setattr(v, "InputTitle", "Campaign type"))
        retry_com(lambda: setattr(v, "InputMessage", INPUT_MESSAGE))
        retry_com(lambda: setattr(v, "ErrorTitle", "Campaign type"))
        retry_com(lambda: setattr(v, "ErrorMessage", ERROR_MESSAGE))
        retry_com(lambda: setattr(v, "ShowInput", True))
        retry_com(lambda: setattr(v, "ShowError", False))
    print("    Dropdowns A2:A10 + validation re-applied")


def process(excel, path: Path):
    print(f"  Opening {path.name}")
    wb = retry_com(lambda: excel.Workbooks.Open(str(path), UpdateLinks=0))
    try:
        name = wb.Name
        edit_vba(wb)
        apply_dropdown_and_validation(wb)
        retry_com(lambda: excel.CalculateFull())
        validation = retry_com(lambda: excel.Run(f"'{name}'!ValidateWorkbookConfiguration"))
        dd = retry_com(lambda: wb.Worksheets("Dropdowns"))
        listed = [retry_com(lambda r=r: dd.Cells(r, 1).Value) for r in range(2, 11)]
        print(f"    Dropdowns A2:A10={listed}; ValidateWorkbookConfiguration={validation!r}")
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
    print("\nCAMPAIGN TYPE CHANGES APPLIED SUCCESSFULLY")
    return 0


if __name__ == "__main__":
    sys.exit(main())
