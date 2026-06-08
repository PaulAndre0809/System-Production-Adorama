"""QA checks for Email & SMS Campaign Tracker.xlsm.

This script is for development/validation only. The workbook does not depend on
this file at runtime.
"""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client as win32


EXPECTED_CAMPAIGN_TYPES = [
    "Promo",
    "Services",
    "Loyalty & PLCC",
    "Newsletters",
    "Events",
    "NPA",
    "Others",
    None,
]

EXPECTED_LAST_REFRESH_FORMULA = '=TEXT(NOW(),"m/d/yyyy h:mm AM/PM")'


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


def normalize_header(value) -> str:
    return "".join(ch for ch in str(value).lower().strip() if ch not in " /\\-")


def column_by_header(list_object, header_name: str):
    wanted = normalize_header(header_name)
    for column in list_object.ListColumns:
        if normalize_header(column.Name) == wanted:
            return column
    return None


def nonempty(value) -> bool:
    return value is not None and str(value).strip() != ""


def assert_zip_integrity(path: Path, checks: list[str]) -> None:
    with zipfile.ZipFile(path) as archive:
        bad_file = archive.testzip()
        if bad_file:
            raise AssertionError(f"Workbook zip integrity failed at {bad_file}")
        if "xl/vbaProject.bin" not in set(archive.namelist()):
            raise AssertionError("Workbook is missing embedded VBA project")
    checks.append("zip integrity and embedded VBA project")


def validate_workbook(path: Path) -> list[str]:
    checks: list[str] = []
    assert_zip_integrity(path, checks)

    temp_dir = Path(tempfile.mkdtemp(prefix="campaign_tracker_qa_"))
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
            excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception:
            pass

        workbook = retry(
            "open workbook",
            lambda: excel.Workbooks.Open(
                str(qa_path), UpdateLinks=0, ReadOnly=False, AddToMru=False
            ),
        )

        validation = retry(
            "embedded validation",
            lambda: excel.Run(macro_name(workbook, "ValidateWorkbookConfiguration")),
        )
        if validation != "OK":
            raise AssertionError(f"Embedded validation failed: {validation}")
        checks.append("embedded ValidateWorkbookConfiguration returned OK")

        dashboard = workbook.Worksheets("Dashboard")
        email_table = workbook.Worksheets("Email Campaigns").ListObjects(
            "EmailCampaignsTable"
        )
        sms_table = workbook.Worksheets("SMS Campaigns").ListObjects(
            "SMSCampaignsTable"
        )
        dashboard_table = dashboard.ListObjects("DashboardWorkTable")

        if dashboard.Range("B3").Formula != EXPECTED_LAST_REFRESH_FORMULA:
            raise AssertionError(f"Unexpected Last Refresh formula: {dashboard.Range('B3').Formula}")
        if dashboard.Range("C3").Value != "Last Edited By":
            raise AssertionError("Dashboard Last Edited field was not removed")
        if dashboard.Range("D3").Formula == "":
            raise AssertionError("Dashboard Last Edited By formula is missing")
        if dashboard.Range("E3").Value is not None or dashboard.Range("F3").Value is not None:
            raise AssertionError("Dashboard still has leftover Last Edited audit fields")
        if dashboard.Range("B3").NumberFormat == "@" or dashboard.Range("D3").NumberFormat == "@":
            raise AssertionError("Dashboard audit formulas are formatted as text")
        retry("calculate dashboard audit cells", lambda: dashboard.Range("B3:D3").Calculate())
        if not isinstance(dashboard.Range("B3").Value, str):
            raise AssertionError("Dashboard Last Refresh is not displaying as text")
        if not isinstance(dashboard.Range("D3").Value, str):
            raise AssertionError("Dashboard Last Edited By is not displaying as text")
        if dashboard.Columns("B").ColumnWidth < 22:
            raise AssertionError("Dashboard Last Refresh column is too narrow")
        if dashboard.Columns("D").ColumnWidth < 22:
            raise AssertionError("Dashboard Last Edited By column is too narrow")
        checks.append("Dashboard refresh and audit user formulas and widths")

        if not email_table.ShowAutoFilter:
            raise AssertionError("Email Campaigns table filter dropdowns are disabled")
        if not sms_table.ShowAutoFilter:
            raise AssertionError("SMS Campaigns table filter dropdowns are disabled")
        checks.append("campaign table filter dropdowns")

        for sheet_index in range(1, workbook.Worksheets.Count + 1):
            sheet = workbook.Worksheets(sheet_index)
            sheet.Activate()
            if excel.ActiveWindow.FreezePanes:
                raise AssertionError(f"Freeze panes still enabled: {sheet.Name}")
            if excel.ActiveWindow.SplitRow != 0 or excel.ActiveWindow.SplitColumn != 0:
                raise AssertionError(f"Split view still enabled: {sheet.Name}")
        checks.append("no freeze panes or split views")

        dropdowns = workbook.Worksheets("Dropdowns")
        actual_types = [dropdowns.Cells(row, 1).Value for row in range(2, 10)]
        if actual_types != EXPECTED_CAMPAIGN_TYPES:
            raise AssertionError(f"Unexpected Campaign Type source list: {actual_types}")
        for table in (email_table, sms_table):
            campaign_type_col = column_by_header(table, "Campaign Type")
            if campaign_type_col is None:
                raise AssertionError(f"Missing Campaign Type column on {table.Name}")
            validation_obj = campaign_type_col.DataBodyRange.Validation
            if validation_obj.Type != 3:
                raise AssertionError(f"Campaign Type is not a list validation on {table.Name}")
            if validation_obj.Formula1 != "=Dropdowns!$A$2:$A$9":
                raise AssertionError(
                    f"Campaign Type source is wrong on {table.Name}: {validation_obj.Formula1}"
                )
            if validation_obj.ShowError:
                raise AssertionError(f"Campaign Type custom values are blocked on {table.Name}")
        checks.append("Campaign Type dropdowns and custom-value behavior")

        for table in (email_table, sms_table):
            send_date = column_by_header(table, "Send Date")
            send_time = column_by_header(table, "Send Time")
            last_updated = column_by_header(table, "Last Updated")
            last_updated_by = column_by_header(table, "Last Updated By")
            if send_date is None or send_time is None or last_updated is None:
                raise AssertionError(f"Missing date/time columns on {table.Name}")
            if send_date.DataBodyRange.NumberFormat.lower() != "mm/dd/yyyy":
                raise AssertionError(f"Send Date format is wrong on {table.Name}")
            if "am/pm" not in str(send_time.DataBodyRange.NumberFormat).lower():
                raise AssertionError(f"Send Time format is not 12-hour on {table.Name}")
            if "am/pm" not in str(last_updated.DataBodyRange.NumberFormat).lower():
                raise AssertionError(f"Last Updated format is not 12-hour on {table.Name}")
            if last_updated_by is None:
                raise AssertionError(f"Missing Last Updated By on {table.Name}")
            last_updated_by.DataBodyRange.Cells(1, 1).Value = "Manual Web User"
            if last_updated_by.DataBodyRange.Cells(1, 1).Value != "Manual Web User":
                raise AssertionError(f"Last Updated By is not manually editable on {table.Name}")
        checks.append("MM/DD/YYYY dates, 12-hour times, and editable audit user cells")

        # Desktop Excel only: verify the embedded Worksheet_Change path stamps the
        # row audit columns when a user-editable campaign cell changes.
        owner_col = column_by_header(email_table, "Owner")
        last_updated_col = column_by_header(email_table, "Last Updated")
        last_updated_by_col = column_by_header(email_table, "Last Updated By")
        if owner_col is None or last_updated_col is None or last_updated_by_col is None:
            raise AssertionError("Missing columns needed for desktop audit stamping test")
        old_timestamp = last_updated_col.DataBodyRange.Cells(1, 1).Value2
        old_owner = owner_col.DataBodyRange.Cells(1, 1).Value
        excel.EnableEvents = True
        owner_col.DataBodyRange.Cells(1, 1).Value = f"{old_owner or 'QA User'} audit test"
        retry("desktop audit event calculate", lambda: dashboard.Range("B3:D3").Calculate())
        excel.EnableEvents = False
        new_timestamp = last_updated_col.DataBodyRange.Cells(1, 1).Value2
        new_user = last_updated_by_col.DataBodyRange.Cells(1, 1).Value
        if new_timestamp is None or (
            old_timestamp is not None and float(new_timestamp) <= float(old_timestamp)
        ):
            raise AssertionError("Desktop Worksheet_Change did not update Last Updated")
        if not nonempty(new_user):
            raise AssertionError("Desktop Worksheet_Change did not update Last Updated By")
        checks.append("desktop edit audit stamping")

        if dashboard_table.ListRows.Count < 150:
            raise AssertionError("Dashboard table is not prepared for native formula rows")
        if not str(dashboard.Range("AA11").Formula2).startswith("=LET("):
            raise AssertionError("Dashboard native helper formula is missing")
        if not dashboard.Columns("AA:AL").Hidden:
            raise AssertionError("Dashboard helper columns AA:AL should be hidden")
        checks.append("Dashboard native formula feed")

        links = retry("external links", lambda: workbook.LinkSources(1))
        if links is not None:
            raise AssertionError(f"External workbook links detected: {links}")
        checks.append("no external workbook links")

        workbook.Close(SaveChanges=False)
    finally:
        if excel is not None:
            excel.Quit()
        shutil.rmtree(temp_dir, ignore_errors=True)

    return checks


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "workbook",
        nargs="?",
        default=r"Production Tracker\Email & SMS Campaign Tracker.xlsm",
        help="Path to the workbook to validate.",
    )
    args = parser.parse_args()

    workbook_path = Path(args.workbook).resolve()
    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 2

    checks = validate_workbook(workbook_path)
    print("QA PASSED")
    for check in checks:
        print(f"- {check}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
