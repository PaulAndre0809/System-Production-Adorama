"""Inspect the current state of Email & SMS Campaign Tracker.xlsm.

Prints sheet names, calendar titles, Current Stage formulas, checkbox column
structure, and other details needed to plan modifications.
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


def macro_name(workbook, proc_name: str) -> str:
    return f"'{workbook.Name}'!{proc_name}"


def main() -> int:
    path = Path(r"Production Tracker\Email & SMS Campaign Tracker.xlsm").resolve()
    if not path.exists():
        print(f"Not found: {path}", file=sys.stderr)
        return 2

    temp_dir = Path(tempfile.mkdtemp(prefix="inspect_"))
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

        workbook = retry(
            "open workbook",
            lambda: excel.Workbooks.Open(str(qa_path), UpdateLinks=0, ReadOnly=True, AddToMru=False),
        )

        # 1. List all sheet names
        print("=" * 60)
        print("SHEET NAMES")
        print("=" * 60)
        for i in range(1, workbook.Worksheets.Count + 1):
            ws = workbook.Worksheets(i)
            print(f"  [{i}] {ws.Name!r}  (tab color: {ws.Tab.Color})")

        # 2. Calendar sheet details — titles
        print("\n" + "=" * 60)
        print("CALENDAR SHEET TITLES (A1 and A2)")
        print("=" * 60)
        months = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December",
        ]
        for month in months:
            try:
                ws = workbook.Worksheets(f"{month} Calendar")
                a1_val = ws.Range("A1").Value
                a1_formula = ws.Range("A1").Formula
                a2_val = ws.Range("A2").Value
                a2_formula = ws.Range("A2").Formula
                b1_val = ws.Range("B1").Value
                b1_formula = ws.Range("B1").Formula
                print(f"  {month} Calendar:")
                print(f"    A1 value={a1_val!r}  formula={a1_formula!r}")
                print(f"    A2 value={a2_val!r}  formula={a2_formula!r}")
                print(f"    B1 value={b1_val!r}  formula={b1_formula!r}")
                # Check the first few rows for header content
                for r in range(1, 6):
                    vals = []
                    for c in range(1, 8):
                        v = ws.Cells(r, c).Value
                        f = ws.Cells(r, c).Formula
                        if v is not None or f:
                            vals.append(f"Col{c}={v!r} (formula={f!r})")
                    if vals:
                        print(f"    Row {r}: {'; '.join(vals)}")
            except Exception as e:
                print(f"  {month} Calendar: NOT FOUND ({e})")

        # 3. Current Stage formula on Email Campaigns
        print("\n" + "=" * 60)
        print("CURRENT STAGE FORMULA - EMAIL")
        print("=" * 60)
        email_ws = workbook.Worksheets("Email Campaigns")
        email_table = email_ws.ListObjects("EmailCampaignsTable")
        for col in email_table.ListColumns:
            if "stage" in col.Name.lower():
                formula = col.DataBodyRange.Cells(1, 1).Formula
                formula2 = col.DataBodyRange.Cells(1, 1).Formula2
                print(f"  Column: {col.Name!r} (index={col.Index})")
                print(f"  Formula: {formula!r}")
                print(f"  Formula2: {formula2!r}")
                # Also print the current value
                val = col.DataBodyRange.Cells(1, 1).Value
                print(f"  Value: {val!r}")
                break

        # 4. Current Stage formula on SMS Campaigns
        print("\n" + "=" * 60)
        print("CURRENT STAGE FORMULA - SMS")
        print("=" * 60)
        sms_ws = workbook.Worksheets("SMS Campaigns")
        sms_table = sms_ws.ListObjects("SMSCampaignsTable")
        for col in sms_table.ListColumns:
            if "stage" in col.Name.lower():
                formula = col.DataBodyRange.Cells(1, 1).Formula
                formula2 = col.DataBodyRange.Cells(1, 1).Formula2
                print(f"  Column: {col.Name!r} (index={col.Index})")
                print(f"  Formula: {formula!r}")
                print(f"  Formula2: {formula2!r}")
                val = col.DataBodyRange.Cells(1, 1).Value
                print(f"  Value: {val!r}")
                break

        # 5. List all column headers for both tables
        print("\n" + "=" * 60)
        print("EMAIL CAMPAIGNS TABLE HEADERS")
        print("=" * 60)
        for col in email_table.ListColumns:
            print(f"  [{col.Index}] {col.Name!r}")

        print("\n" + "=" * 60)
        print("SMS CAMPAIGNS TABLE HEADERS")
        print("=" * 60)
        for col in sms_table.ListColumns:
            print(f"  [{col.Index}] {col.Name!r}")

        # 6. Check existing data row count
        print("\n" + "=" * 60)
        print("DATA ROW COUNTS")
        print("=" * 60)
        print(f"  Email rows: {email_table.ListRows.Count}")
        print(f"  SMS rows: {sms_table.ListRows.Count}")

        # 7. Inspect Est. Audience and Delivered column indices
        print("\n" + "=" * 60)
        print("KEY COLUMN INDICES")
        print("=" * 60)
        for table_name, table in [("Email", email_table), ("SMS", sms_table)]:
            for col in table.ListColumns:
                if col.Name in ("Est. Audience", "Delivered", "Current Stage"):
                    print(f"  {table_name}: {col.Name!r} = col {col.Index}")

        # 8. Check if there are any named ranges related to calendars
        print("\n" + "=" * 60)
        print("NAMED RANGES")
        print("=" * 60)
        for name in workbook.Names:
            print(f"  {name.Name!r} => {name.RefersTo!r}")

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
