"""Apply workbook modifications for Email & SMS Campaign Tracker.

Modifications:
1. Replace Calendar sheet formulas in A1 with static text.
2. Rename Calendar sheet tabs to include '2026' prefix.
3. Update Current Stage formula to show 'Completed' when checkboxes+audience are filled.
4. Export VBA code, update calendar sheet references to '2026 {Month} Calendar', and re-import.
"""
from __future__ import annotations

import shutil
import sys
import tempfile
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client as win32

MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]


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


def main() -> int:
    path = Path(r"Production Tracker\Email & SMS Campaign Tracker.xlsm").resolve()
    if not path.exists():
        print(f"Not found: {path}", file=sys.stderr)
        return 2

    # Use backup path as source if it exists so we can run repeatedly
    backup_path = path.with_name(path.stem + "_backup" + path.suffix)
    if backup_path.exists():
        shutil.copy2(backup_path, path)
        print("Restored from backup for clean run.")
    else:
        shutil.copy2(path, backup_path)
        print(f"Backup created at: {backup_path}")

    temp_dir = Path(tempfile.mkdtemp(prefix="mod_"))

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

        workbook = retry(
            "open workbook",
            lambda: excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False, AddToMru=False),
        )

        # 1 & 2: Update Calendar Titles and Tab Names
        for month in MONTHS:
            old_name = f"{month} Calendar"
            new_name = f"2026 {month} Calendar"
            try:
                ws = workbook.Worksheets(old_name)
                ws.Name = new_name
                ws.Range("A1").Formula = ""
                ws.Range("A1").Value = f"{month} 2026 Campaign Calendar"
                print(f"Updated {old_name} -> {new_name}")
            except Exception as e:
                print(f"Warning: Could not update {old_name} - {e}")

        # 3. Update Current Stage Formulas
        email_ws = workbook.Worksheets("Email Campaigns")
        email_table = email_ws.ListObjects("EmailCampaignsTable")
        sms_ws = workbook.Worksheets("SMS Campaigns")
        sms_table = sms_ws.ListObjects("SMSCampaignsTable")

        # The new formula uses AND to check all checkboxes and Est. Audience > 0.
        # Email checkboxes: Campaign Name and UTM Parameter (Source Code), Creative Brief, SL & PH, SKUs, In-Design, Build, QA, Route, Approval, Segments
        email_stage_formula = (
            '=IF([@[Campaign Name]]="","",'
            'LET(checked,TEXTJOIN(", ",TRUE,'
            'IF([@[Campaign Name and UTM Parameter (Source Code)]],"Campaign Name and UTM Parameter (Source Code)",""),'
            'IF([@[Creative Brief, SL & PH]],"Creative Brief, SL & PH",""),'
            'IF([@SKUs],"SKUs",""),'
            'IF([@[In-Design]],"In-Design",""),'
            'IF([@[Build, QA]],"Build, QA",""),'
            'IF([@Route],"Route",""),'
            'IF([@Approval],"Approval",""),'
            'IF([@Segments],"Segments","")),'
            'allDone,AND('
            '[@[Campaign Name and UTM Parameter (Source Code)]],'
            '[@[Creative Brief, SL & PH]],'
            '[@SKUs],'
            '[@[In-Design]],'
            '[@[Build, QA]],'
            '[@Route],'
            '[@Approval],'
            '[@Segments]),'
            'IF(AND(allDone, [@[Est. Audience]]<>"", [@[Est. Audience]]>0),"Completed",'
            'IF(checked="","No checklist items checked","Done: "&checked))))'
        )

        for col in email_table.ListColumns:
            if "stage" in col.Name.lower():
                try:
                    col.DataBodyRange.Cells(1, 1).Formula2 = email_stage_formula
                except Exception:
                    col.DataBodyRange.Cells(1, 1).Formula = email_stage_formula
                print("Updated Email Campaigns Stage formula")
                break

        # SMS checkboxes: Send SMS Options, Send Test, Approval, Segments
        sms_stage_formula = (
            '=IF([@[Campaign Name]]="","",'
            'LET(checked,TEXTJOIN(", ",TRUE,'
            'IF([@[Send SMS Options]],"Send SMS Options",""),'
            'IF([@[Send Test]],"Send Test",""),'
            'IF([@Approval],"Approval",""),'
            'IF([@Segments],"Segments","")),'
            'allDone,AND('
            '[@[Send SMS Options]],'
            '[@[Send Test]],'
            '[@Approval],'
            '[@Segments]),'
            'IF(AND(allDone, [@[Est. Audience]]<>"", [@[Est. Audience]]>0),"Completed",'
            'IF(checked="","No checklist items checked","Done: "&checked))))'
        )

        for col in sms_table.ListColumns:
            if "stage" in col.Name.lower():
                try:
                    col.DataBodyRange.Cells(1, 1).Formula2 = sms_stage_formula
                except Exception:
                    col.DataBodyRange.Cells(1, 1).Formula = sms_stage_formula
                print("Updated SMS Campaigns Stage formula")
                break

        # 4. Update VBA Macro References
        # The VBA code has references like: `MonthName(monthNumber) & " Calendar"`
        # We'll export the modules, search and replace, and import them back.
        modules_updated = 0
        for i in range(1, workbook.VBProject.VBComponents.Count + 1):
            comp = retry(f"comp {i}", lambda i=i: workbook.VBProject.VBComponents(i))
            name = retry(f"name {i}", lambda: comp.Name)
            count = retry(f"count {i}", lambda: comp.CodeModule.CountOfLines)
            
            if count > 0:
                code = retry(f"code {name}", lambda: comp.CodeModule.Lines(1, count))
                
                # Several replacements needed:
                # 'MonthName(monthNumber) & " Calendar"' -> '"2026 " & MonthName(monthNumber) & " Calendar"'
                # '"'" & MonthName(previousMonth) & " Calendar'!A1"' -> '"'2026 " & MonthName(previousMonth) & " Calendar'!A1"'
                # '"'" & MonthName(nextMonth) & " Calendar'!A1"' -> '"'2026 " & MonthName(nextMonth) & " Calendar'!A1"'
                # '"'" & MonthName(monthNumber) & " Calendar'!A1"' -> '"'2026 " & MonthName(monthNumber) & " Calendar'!A1"'
                # 'InStr(1, ws.Name, " Calendar"' -> 'InStr(1, ws.Name, "2026 ", vbTextCompare) > 0 And InStr(1, ws.Name, " Calendar"' (this is safer: just look for Calendar but we don't strictly need to change it if it just looks for ' Calendar' as a suffix, actually ' Calendar' is still at the end so InStr(1, ws.Name, " Calendar") still works!)
                # 'MonthName(monthNumber) & " Calendar"'
                
                original_code = code
                
                code = code.replace(
                    'MonthName(monthNumber) & " Calendar"',
                    '"2026 " & MonthName(monthNumber) & " Calendar"'
                )
                code = code.replace(
                    'MonthName(previousMonth) & " Calendar',
                    '2026 " & MonthName(previousMonth) & " Calendar'
                )
                code = code.replace(
                    'MonthName(nextMonth) & " Calendar',
                    '2026 " & MonthName(nextMonth) & " Calendar'
                )
                code = code.replace(
                    'targetName = MonthName(monthNumber) & " Calendar"',
                    'targetName = "2026 " & MonthName(monthNumber) & " Calendar"'
                )
                
                # Fix SMS Headers validation array in VBA
                # The original code expects Jira Link and ClickUp Link in SMS, but they were removed for Proof of Schedule
                code = code.replace(
                    '"Segments", _\r\n\r\n        "Jira Link", _\r\n\r\n        "ClickUp Link", _',
                    '"Segments", _\r\n\r\n        "Proof of Schedule", _'
                )
                code = code.replace(
                    '"Segments", _\n\n        "Jira Link", _\n\n        "ClickUp Link", _',
                    '"Segments", _\n\n        "Proof of Schedule", _'
                )
                
                # Also replace the 'ws.Range("A1").Formula =' part
                
                code = code.replace(
                    'ws.Range("A1").Formula = "=TEXT(DATE(YEAR(TODAY())," & monthNumber & ",1),""""mmmm yyyy"""")&"""" Campaign Calendar""""',
                    'ws.Range("A1").value = MonthName(monthNumber) & " 2026 Campaign Calendar"\n    ws.Range("A1").Formula = ""'
                )
                
                # Check for `MonthName(monthNumber) & " Calendar"` where it might have been missed
                # For `InStr(1, ws.Name, " Calendar"`, it's still fine.

                if original_code != code:
                    # Update module
                    retry(f"delete {name}", lambda: comp.CodeModule.DeleteLines(1, count))
                    retry(f"insert {name}", lambda: comp.CodeModule.AddFromString(code))
                    print(f"Updated VBA code in module: {name}")
                    modules_updated += 1

        print(f"Total VBA modules updated: {modules_updated}")

        workbook.Save()

        # 5. Create Template Copy
        template_path = Path(r"Production Tracker\Email & SMS Campaign Tracker Template.xlsm").resolve()
        
        # Clear data from Email Campaigns
        if email_table.ListRows.Count > 0:
            email_table.DataBodyRange.Rows(f"2:{email_table.ListRows.Count}").Delete()
            # Clear first row values but keep formulas
            for col_idx in range(1, email_table.ListColumns.Count + 1):
                cell = email_table.DataBodyRange.Cells(1, col_idx)
                if not cell.HasFormula:
                    cell.Value = ""
        
        # Clear data from SMS Campaigns
        if sms_table.ListRows.Count > 0:
            sms_table.DataBodyRange.Rows(f"2:{sms_table.ListRows.Count}").Delete()
            for col_idx in range(1, sms_table.ListColumns.Count + 1):
                cell = sms_table.DataBodyRange.Cells(1, col_idx)
                if not cell.HasFormula:
                    cell.Value = ""
                    
        # Clear Automation Log if it has data
        log_ws = workbook.Worksheets("Automation Log")
        log_table = log_ws.ListObjects("AutomationLogTable")
        if log_table.ListRows.Count > 0:
            log_table.DataBodyRange.Rows(f"2:{log_table.ListRows.Count}").Delete()
            for col_idx in range(1, log_table.ListColumns.Count + 1):
                cell = log_table.DataBodyRange.Cells(1, col_idx)
                if not cell.HasFormula:
                    cell.Value = ""

        # Make sure current stage formulas persist
        workbook.SaveAs(str(template_path))
        print(f"Template created at: {template_path}")

        workbook.Close(SaveChanges=False)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        shutil.rmtree(temp_dir, ignore_errors=True)

    return 0

if __name__ == "__main__":
    raise SystemExit(main())
