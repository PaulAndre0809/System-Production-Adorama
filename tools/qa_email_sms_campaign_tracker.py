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
import datetime as dt

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
EXPECTED_SEND_DATE_FORMAT = "dddd, mmmm d, yyyy"
EXPECTED_SEND_TIME_FORMAT = "h:mm am/pm"
NOTES_PASSWORD = "Adorama@042026_"  # verified against the Notes sheet SHA-512 protection hash


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


def range_has_content(cell_range) -> bool:
    """Return True when any cell in a multi-cell COM range has a value."""
    values = cell_range.Value
    if values is None:
        return False
    if not isinstance(values, tuple):
        return nonempty(values)
    return any(
        nonempty(value)
        for row in values
        for value in (row if isinstance(row, tuple) else (row,))
    )


def assert_valid_vba_continuations(workbook, checks: list[str]) -> None:
    """Detect the blank-after-underscore corruption that breaks VBA compile."""
    malformed: list[str] = []
    for component_index in range(1, workbook.VBProject.VBComponents.Count + 1):
        component = workbook.VBProject.VBComponents(component_index)
        code_module = component.CodeModule
        if not code_module.CountOfLines:
            continue
        code = code_module.Lines(1, code_module.CountOfLines)
        lines = code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        for index, line in enumerate(lines[:-1]):
            if line.rstrip().endswith("_") and not lines[index + 1].strip():
                malformed.append(f"{component.Name}:{index + 1}")
    if malformed:
        raise AssertionError(
            "Malformed VBA continuations found: " + ", ".join(malformed[:20])
        )
    checks.append("VBA continuation syntax is structurally valid")


def excel_date(value) -> dt.date | None:
    """Convert an Excel date value to a Python date for KPI verification."""
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)):
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(value))
    return None


def campaign_kpi_counts(table) -> dict[str, int]:
    """Calculate expected KPI values directly from one campaign table."""
    headers = {
        str(table.ListColumns(index).Name): index - 1
        for index in range(1, table.ListColumns.Count + 1)
    }
    values = table.DataBodyRange.Value2
    rows = values if isinstance(values, tuple) else ((values,),)
    counts = {"active": 0, "today": 0, "approval": 0, "sent": 0}

    for row in rows:
        campaign = row[headers["Campaign Name"]]
        if not nonempty(campaign):
            continue
        notes = str(row[headers["Notes"]] or "").strip().lower()
        stage = str(row[headers["Current Stage"]] or "").strip().lower()
        cancelled = (
            notes in {"cancelled", "canceled"}
            or stage in {"cancelled", "canceled"}
        )
        delivered_value = row[headers["Delivered"]]
        try:
            delivered = float(delivered_value or 0)
        except (TypeError, ValueError):
            delivered = 0

        if delivered > 0 and not cancelled:
            counts["sent"] += 1
        elif not cancelled:
            counts["active"] += 1
        if not cancelled and excel_date(row[headers["Send Date"]]) == dt.date.today():
            counts["today"] += 1
        if not cancelled and not bool(row[headers["Approval"]]):
            counts["approval"] += 1
    return counts


def assert_dashboard_kpis(
    dashboard,
    email_table,
    sms_table,
    checks: list[str],
) -> None:
    """Verify KPI formulas and displayed values against source-table data."""
    for address in ("C5", "E5", "G5", "I5", "K5"):
        formula = str(dashboard.Range(address).Formula2)
        if "@Delivered" in formula or "--@Email" in formula or "--@SMS" in formula:
            raise AssertionError(
                f"Dashboard KPI {address} still uses implicit intersection"
            )

    dashboard.Range("A5:K5").Calculate()
    email = campaign_kpi_counts(email_table)
    sms = campaign_kpi_counts(sms_table)
    expected = {
        "A5": email["active"] + sms["active"],
        "C5": email["today"] + sms["today"],
        "E5": email["active"],
        "G5": sms["active"],
        "I5": email["approval"] + sms["approval"],
        "K5": email["sent"] + sms["sent"],
    }
    for address, expected_value in expected.items():
        actual = int(dashboard.Range(address).Value or 0)
        if actual != expected_value:
            raise AssertionError(
                f"Dashboard KPI {address} is {actual}; expected {expected_value}"
            )
    for footer_address, source_address in {
        "C162": "C5",
        "C163": "E5",
        "C164": "G5",
        "C165": "I5",
        "C166": "K5",
    }.items():
        if str(dashboard.Range(footer_address).Formula2) != f"={source_address}":
            raise AssertionError(
                f"Dashboard footer {footer_address} is not linked to {source_address}"
            )
    checks.append("Dashboard KPI formulas use expanding references and match source data")


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
        assert_valid_vba_continuations(workbook, checks)

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

        calendar_sheets = [
            workbook.Worksheets(index).Name
            for index in range(1, workbook.Worksheets.Count + 1)
            if "calendar" in str(workbook.Worksheets(index).Name).lower()
        ]
        if calendar_sheets:
            raise AssertionError(f"Retired Calendar sheets still exist: {calendar_sheets}")
        comparison_tables = [
            dashboard.ListObjects(index).Name
            for index in range(1, dashboard.ListObjects.Count + 1)
            if "deliveredcomparison"
            in str(dashboard.ListObjects(index).Name).lower()
        ]
        if comparison_tables:
            raise AssertionError(
                f"Retired comparison tables still exist: {comparison_tables}"
            )
        comparison_charts = [
            dashboard.ChartObjects(index).Name
            for index in range(1, dashboard.ChartObjects().Count + 1)
            if "delivered" in str(dashboard.ChartObjects(index).Name).lower()
            and "comparison" in str(dashboard.ChartObjects(index).Name).lower()
        ]
        if comparison_charts:
            raise AssertionError(
                f"Retired comparison charts still exist: {comparison_charts}"
            )
        if range_has_content(dashboard.Range("N2:S44")):
            raise AssertionError("Retired Dashboard comparison cells are not empty")
        if range_has_content(
            dashboard.Range("A7:L7")
        ) or dashboard.Range("A7:L7").Hyperlinks.Count:
            raise AssertionError("Retired Dashboard calendar links are not empty")
        notes = workbook.Worksheets("Notes - Instructions")
        if notes.Range("A1").Value != "Detailed Notes and Instructions":
            raise AssertionError("Notes - Instructions title is outdated")
        notes_text = str(notes.UsedRange.Value)
        if "intentionally removed" not in notes_text:
            raise AssertionError("Notes do not document the retired features")
        if "Wednesday, June 10, 2026" not in notes_text:
            raise AssertionError("Notes do not document the Send Date format")
        if "STO or Local Timezone" not in notes_text:
            raise AssertionError("Notes do not document allowed Send Time text")
        if "Timed Link Labels" not in notes_text:
            raise AssertionError("Notes do not document seven-day link labels")
        if "Cancelled Campaigns" not in notes_text:
            raise AssertionError("Notes do not document Dashboard cancellation filtering")
        checks.append("Calendar sheets and weekly Dashboard comparisons are retired")

        # Read via retry: the preceding pure-Python checks give Excel idle time to
        # start a volatile NOW()/TODAY() background recalc, which rejects bare COM
        # reads ("call rejected by callee") on a busy machine.
        last_refresh = retry("read Last Refresh formula", lambda: dashboard.Range("B3").Formula)
        if last_refresh != EXPECTED_LAST_REFRESH_FORMULA:
            raise AssertionError(f"Unexpected Last Refresh formula: {last_refresh}")
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
        assert_dashboard_kpis(dashboard, email_table, sms_table, checks)

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
            if (
                str(send_date.DataBodyRange.NumberFormat).lower()
                != EXPECTED_SEND_DATE_FORMAT
            ):
                raise AssertionError(f"Send Date format is wrong on {table.Name}")
            if (
                str(send_time.DataBodyRange.NumberFormat).lower()
                != EXPECTED_SEND_TIME_FORMAT
            ):
                raise AssertionError(f"Send Time format is not 12-hour on {table.Name}")
            if "am/pm" not in str(last_updated.DataBodyRange.NumberFormat).lower():
                raise AssertionError(f"Last Updated format is not 12-hour on {table.Name}")
            if last_updated_by is None:
                raise AssertionError(f"Missing Last Updated By on {table.Name}")
            last_updated_by.DataBodyRange.Cells(1, 1).Value = "Manual Web User"
            if last_updated_by.DataBodyRange.Cells(1, 1).Value != "Manual Web User":
                raise AssertionError(f"Last Updated By is not manually editable on {table.Name}")
            date_cell = send_date.DataBodyRange.Cells(1, 1)
            time_cell = send_time.DataBodyRange.Cells(1, 1)
            old_date = date_cell.Value2
            old_time = time_cell.Value2
            date_cell.Value2 = 46183  # June 10, 2026
            if str(date_cell.Text) != "Wednesday, June 10, 2026":
                raise AssertionError(
                    f"Send Date display is wrong on {table.Name}: {date_cell.Text}"
                )
            for value in ("STO", "Local Timezone"):
                time_cell.Value2 = value
                if time_cell.Value2 != value:
                    raise AssertionError(
                        f"Send Time rejected {value!r} on {table.Name}"
                    )
            time_cell.Value2 = 10 / 24
            if str(time_cell.Text).upper() != "10:00 AM":
                raise AssertionError(
                    f"Send Time display is wrong on {table.Name}: {time_cell.Text}"
                )
            date_cell.Value2 = old_date
            time_cell.Value2 = old_time

        notes.Unprotect(Password=NOTES_PASSWORD)
        if notes.ProtectContents:
            raise AssertionError("Notes password did not unlock the sheet")
        notes.Protect(
            Password=NOTES_PASSWORD,
            UserInterfaceOnly=True,
            AllowFiltering=True,
        )
        checks.append(
            "long-date display, permissive 12-hour times, and verified Notes password"
        )

        # Desktop Excel only: verify the embedded Worksheet_Change path stamps the
        # row audit columns when a user-editable campaign cell changes.
        owner_col = column_by_header(email_table, "Owner")
        last_updated_col = column_by_header(email_table, "Last Updated")
        last_updated_by_col = column_by_header(email_table, "Last Updated By")
        if owner_col is None or last_updated_col is None or last_updated_by_col is None:
            raise AssertionError("Missing columns needed for desktop audit stamping test")
        old_timestamp = last_updated_col.DataBodyRange.Cells(1, 1).Value2
        old_owner = owner_col.DataBodyRange.Cells(1, 1).Value
        event_date = column_by_header(email_table, "Send Date").DataBodyRange.Cells(1, 1)
        event_time = column_by_header(email_table, "Send Time").DataBodyRange.Cells(1, 1)
        old_event_date = event_date.Value2
        old_event_time = event_time.Value2
        excel.EnableEvents = True

        event_date.NumberFormat = "General"
        event_date.Value2 = 46183
        if (
            str(event_date.NumberFormat).lower() != EXPECTED_SEND_DATE_FORMAT
            or str(event_date.Text) != "Wednesday, June 10, 2026"
        ):
            raise AssertionError("Desktop date edit did not restore long-date formatting")

        event_time.NumberFormat = "General"
        for value in ("STO", "Local Timezone"):
            event_time.Value2 = value
            if (
                event_time.Value2 != value
                or str(event_time.NumberFormat).lower() != EXPECTED_SEND_TIME_FORMAT
            ):
                raise AssertionError(
                    f"Desktop time edit did not preserve {value!r}"
                )
        event_time.Value2 = 10 / 24
        if str(event_time.Text).upper() != "10:00 AM":
            raise AssertionError("Desktop numeric time edit is not 12-hour formatted")

        event_date.Value2 = old_event_date
        event_time.Value2 = old_event_time
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
        checks.append("desktop date/time edit repair and audit stamping")

        if dashboard_table.ListRows.Count < 150:
            raise AssertionError("Dashboard table is not prepared for native formula rows")
        if not str(dashboard.Range("AA11").Formula2).startswith("=LET("):
            raise AssertionError("Dashboard native helper formula is missing")
        dashboard_formula = str(dashboard.Range("AA11").Formula2).lower()
        if "cancelled" not in dashboard_formula or "[current stage]" not in dashboard_formula:
            raise AssertionError("Dashboard helper does not exclude cancelled campaigns")
        if not dashboard.Columns("AA:AL").Hidden:
            raise AssertionError("Dashboard helper columns AA:AL should be hidden")
        checks.append("Dashboard native formula feed excludes cancelled campaigns")

        for table, configurations in (
            (
                email_table,
                (
                    ("Jira Link", "JIRA"),
                    ("ClickUp Link", "ClickUp"),
                    ("Bluecore/Attentive Link", "Bluecore/Attentive"),
                ),
            ),
            (
                sms_table,
                (
                    ("Proof of Schedule", "Proof of Schedule"),
                    ("Bluecore/Attentive Link", "Bluecore/Attentive"),
                ),
            ),
        ):
            for header, display_name in configurations:
                link_cell = column_by_header(table, header).DataBodyRange.Cells(1, 1)
                if not link_cell.HasFormula:
                    continue
                formula = str(link_cell.Formula2)
                if (
                    "HYPERLINK(" not in formula.upper()
                    or "NOW()" not in formula.upper()
                    or "+7" not in formula
                    or display_name not in formula
                ):
                    raise AssertionError(
                        f"Timed hyperlink formula is wrong on {table.Name}[{header}]"
                    )
        checks.append("supported link columns use native seven-day hyperlink formulas")

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
