from __future__ import annotations

import argparse
import gc
import os
import shutil
import sys
import time
import uuid
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Callable
from zipfile import ZipFile

import pythoncom
import pywintypes
import win32com.client


BASE_NAME = "Email & SMS Campaign Tracker"
WORKBOOK_NAME = f"{BASE_NAME}.xlsm"
MODULE = "modEmailProductionTracker"
EMAIL_SHEET = "Email Campaigns"
SMS_SHEET = "SMS Campaigns"
EMAIL_TABLE = "EmailCampaignsTable"
SMS_TABLE = "SMSCampaignsTable"
DASHBOARD_TABLE = "DashboardWorkTable"
COMPARISON_TABLE = "DeliveredComparisonTable"
EMAIL_CHECKLIST = (
    "Campaign Name and UTM Parameter (Source Code)",
    "Creative Brief, SL & PH",
    "SKUs",
    "In-Design",
    "Build, QA",
    "Route",
    "Approval",
    "Segments",
)
SMS_CHECKLIST = ("Send SMS Options", "Send Test", "Approval", "Segments")
COMMON_PREFIX = (
    "Send Date",
    "Send Time",
    "Campaign Name",
    "Campaign Type",
    "Current Stage",
    "Owner",
)
COMMON_SUFFIX = (
    "Jira Link",
    "ClickUp Link",
    "Bluecore Link",
    "Est. Audience",
    "Delivered",
    "Last Updated",
    "Last Updated By",
)
EMAIL_HEADERS = (*COMMON_PREFIX, *EMAIL_CHECKLIST, *COMMON_SUFFIX)
SMS_HEADERS = (*COMMON_PREFIX, *SMS_CHECKLIST, *COMMON_SUFFIX)
RPC_E_CALL_REJECTED = -2147418111
XL_CALCULATION_AUTOMATIC = -4105
QA_EMAIL = f"QA EMAIL {uuid.uuid4().hex[:8].upper()}"
QA_SMS = f"QA SMS {uuid.uuid4().hex[:8].upper()}"
QA_LAST_WEEK = f"QA LAST WEEK {uuid.uuid4().hex[:8].upper()}"


def com_retry(
    operation: Callable[[], Any],
    attempts: int = 40,
    delay: float = 0.25,
) -> Any:
    for attempt in range(attempts):
        try:
            return operation()
        except pywintypes.com_error as exc:
            if exc.hresult != RPC_E_CALL_REJECTED or attempt == attempts - 1:
                raise
            time.sleep(delay)
    raise RuntimeError("Excel COM operation did not complete.")


def run_macro(excel: Any, workbook_name: str, procedure: str, *args: Any) -> Any:
    name = f"'{workbook_name}'!{MODULE}.{procedure}"
    return com_retry(lambda: excel.Run(name, *args))


def rgb(red: int, green: int, blue: int) -> int:
    return red + green * 256 + blue * 65536


def flatten(values: Any) -> list[Any]:
    if isinstance(values, tuple):
        output: list[Any] = []
        for value in values:
            output.extend(flatten(value))
        return output
    return [values]


def has_validation(cell_range: Any) -> bool:
    try:
        return com_retry(lambda: cell_range.Validation.Type) not in (None, 0)
    except pywintypes.com_error:
        return False


def table_headers(table: Any) -> tuple[str, ...]:
    return tuple(
        table.ListColumns(index).Name
        for index in range(1, table.ListColumns.Count + 1)
    )


def cell_for(table: Any, row_index: int, header: str) -> Any:
    column_index = table.ListColumns(header).Index
    return table.DataBodyRange.Rows(row_index).Cells(1, column_index)


def find_blank_row(table: Any) -> int:
    campaigns = table.ListColumns("Campaign Name").DataBodyRange
    for index in range(table.ListRows.Count, 0, -1):
        value = campaigns.Cells(index, 1).Value2
        if value is None or not str(value).strip():
            return index
    table.ListRows.Add()
    return table.ListRows.Count


def excel_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
    raise ValueError(f"Not an Excel date: {value!r}")


def formula_errors(workbook: Any) -> list[str]:
    errors: list[str] = []
    for sheet in workbook.Worksheets:
        for formula in flatten(sheet.UsedRange.Formula):
            if isinstance(formula, str) and "#REF!" in formula.upper():
                errors.append(sheet.Name)
                break
    for name in workbook.Names:
        if "#REF!" in str(name.RefersTo).upper():
            errors.append(f"Defined name: {name.Name}")
    return errors


def checklist_checks(table: Any, headers: tuple[str, ...]) -> tuple[bool, bool, bool]:
    native = True
    no_dropdowns = True
    booleans = True
    for header in headers:
        data_range = table.ListColumns(header).DataBodyRange
        native = native and data_range.CellControl.Type == 2
        no_dropdowns = no_dropdowns and not has_validation(data_range)
        booleans = booleans and all(
            value is None or isinstance(value, bool)
            for value in flatten(data_range.Value2)
        )
    return native, no_dropdowns, booleans


def structural_checks(excel: Any, workbook: Any) -> dict[str, bool]:
    sheet_names = [sheet.Name for sheet in workbook.Worksheets]
    email = workbook.Worksheets(EMAIL_SHEET)
    sms = workbook.Worksheets(SMS_SHEET)
    email_table = email.ListObjects(EMAIL_TABLE)
    sms_table = sms.ListObjects(SMS_TABLE)
    dashboard = workbook.Worksheets("Dashboard")
    dashboard_table = dashboard.ListObjects(DASHBOARD_TABLE)
    comparison = dashboard.ListObjects(COMPARISON_TABLE)

    email_native, email_no_dropdowns, email_booleans = checklist_checks(
        email_table, EMAIL_CHECKLIST
    )
    sms_native, sms_no_dropdowns, sms_booleans = checklist_checks(
        sms_table, SMS_CHECKLIST
    )

    calendar_ok = True
    calendar_colors_ok = True
    for month_number in range(1, 13):
        sheet_name = datetime(2000, month_number, 1).strftime("%B Calendar")
        calendar = workbook.Worksheets(sheet_name)
        formula = str(calendar.Range("A6").Formula)
        calendar_ok = (
            calendar_ok
            and EMAIL_TABLE.lower() in formula.lower()
            and SMS_TABLE.lower() in formula.lower()
        )
        expected = (
            rgb(0, 176, 80)
            if month_number == date.today().month
            else rgb(91, 155, 213)
        )
        calendar_colors_ok = (
            calendar_colors_ok and int(calendar.Tab.Color) == expected
        )

    dashboard_headers = tuple(
        dashboard_table.ListColumns(index).Name
        for index in range(1, dashboard_table.ListColumns.Count + 1)
    )
    links = com_retry(lambda: workbook.LinkSources(1))
    return {
        "email_sheet_and_table": EMAIL_SHEET in sheet_names,
        "sms_sheet_and_table": SMS_SHEET in sheet_names,
        "production_inventory_removed": "Production Inventory" not in sheet_names,
        "email_headers": table_headers(email_table) == EMAIL_HEADERS,
        "sms_headers": table_headers(sms_table) == SMS_HEADERS,
        "email_checkboxes_g_to_n": [
            email_table.ListColumns(h).Index for h in EMAIL_CHECKLIST
        ]
        == list(range(7, 15)),
        "sms_checkboxes_g_to_j": [
            sms_table.ListColumns(h).Index for h in SMS_CHECKLIST
        ]
        == list(range(7, 11)),
        "native_email_checkboxes": email_native,
        "native_sms_checkboxes": sms_native,
        "checkbox_dropdowns_removed": email_no_dropdowns and sms_no_dropdowns,
        "checkbox_values_boolean": email_booleans and sms_booleans,
        "owner_plain_text": (
            not has_validation(email_table.ListColumns("Owner").DataBodyRange)
            and not has_validation(sms_table.ListColumns("Owner").DataBodyRange)
        ),
        "email_stage_formula": bool(
            email_table.ListColumns("Current Stage").DataBodyRange.HasFormula
        ),
        "sms_stage_formula": bool(
            sms_table.ListColumns("Current Stage").DataBodyRange.HasFormula
        ),
        "audit_columns_positioned": (
            email_table.ListColumns("Last Updated By").Index
            == email_table.ListColumns("Last Updated").Index + 1
            and sms_table.ListColumns("Last Updated By").Index
            == sms_table.ListColumns("Last Updated").Index + 1
        ),
        "calendar_combines_email_sms": calendar_ok,
        "calendar_tab_colors": calendar_colors_ok,
        "dashboard_combined_headers": dashboard_headers
        == (
            "Send Date",
            "Time",
            "Channel",
            "Campaign",
            "Type",
            "Stage",
            "Owner",
            "Approval",
            "Segments",
            "Jira",
            "ClickUp",
            "Bluecore",
        ),
        "delivered_comparison_table": (
            comparison.ListColumns.Count == 5
            and comparison.ListRows.Count == 2
        ),
        "forced_full_calculation_disabled": not bool(workbook.ForceFullCalculation),
        "no_external_links": not links,
        "no_broken_references": not formula_errors(workbook),
    }


def set_common_row(
    table: Any,
    row_index: int,
    campaign: str,
    send_date: datetime,
) -> None:
    cell_for(table, row_index, "Send Date").Value = send_date
    cell_for(table, row_index, "Send Time").Value = 0.5
    cell_for(table, row_index, "Campaign Name").Value2 = campaign
    cell_for(table, row_index, "Campaign Type").Value2 = "QA"
    cell_for(table, row_index, "Owner").Value2 = "QA User"


def complete_links(table: Any, row_index: int) -> None:
    cell_for(table, row_index, "Jira Link").Value2 = "https://example.com/jira"
    cell_for(table, row_index, "ClickUp Link").Value2 = "https://example.com/clickup"
    cell_for(table, row_index, "Bluecore Link").Value2 = "https://example.com/bluecore"
    cell_for(table, row_index, "Est. Audience").Value2 = 1000


def sum_delivered_for_week(table: Any, week_start: date) -> float:
    total = 0.0
    week_end = week_start + timedelta(days=6)
    send_dates = flatten(table.ListColumns("Send Date").DataBodyRange.Value)
    delivered_values = flatten(
        table.ListColumns("Delivered").DataBodyRange.Value2
    )
    for send_value, delivered in zip(send_dates, delivered_values):
        try:
            send_day = excel_date(send_value)
        except ValueError:
            continue
        if week_start <= send_day <= week_end and isinstance(delivered, (int, float)):
            total += float(delivered)
    return total


def behavior_checks(excel: Any, workbook: Any) -> dict[str, bool]:
    workbook_name = workbook.Name
    email = workbook.Worksheets(EMAIL_SHEET)
    sms = workbook.Worksheets(SMS_SHEET)
    email_table = email.ListObjects(EMAIL_TABLE)
    sms_table = sms.ListObjects(SMS_TABLE)

    email_row = find_blank_row(email_table)
    sms_row = find_blank_row(sms_table)
    last_week_row = find_blank_row(email_table)
    if last_week_row == email_row:
        email_table.ListRows.Add()
        last_week_row = email_table.ListRows.Count

    set_common_row(email_table, email_row, QA_EMAIL, datetime.now())
    for header in EMAIL_CHECKLIST:
        cell_for(email_table, email_row, header).Value2 = False
    email_stage = cell_for(email_table, email_row, "Current Stage")
    email_stage.Calculate()
    email_initial_stage = str(email_stage.Value2 or "")
    cell_for(
        email_table,
        email_row,
        "Campaign Name and UTM Parameter (Source Code)",
    ).Value2 = True
    email_stage.Calculate()
    email_second_stage = str(email_stage.Value2 or "")

    for header in EMAIL_CHECKLIST:
        cell_for(email_table, email_row, header).Value2 = True
    complete_links(email_table, email_row)
    email_stage.Calculate()
    email_scheduled_stage = str(email_stage.Value2 or "")

    run_macro(
        excel,
        workbook_name,
        "HandleCampaignChange",
        email,
        cell_for(email_table, email_row, "Owner"),
    )
    email_timestamp = cell_for(email_table, email_row, "Last Updated").Value2
    email_user = str(
        cell_for(email_table, email_row, "Last Updated By").Value2 or ""
    ).strip()
    cell_for(email_table, email_row, "Delivered").Value2 = 250
    email_stage.Calculate()
    email_sent_stage = str(email_stage.Value2 or "")

    current_monday = date.today() - timedelta(days=date.today().weekday())
    previous_monday = current_monday - timedelta(days=7)
    set_common_row(
        email_table,
        last_week_row,
        QA_LAST_WEEK,
        datetime.combine(previous_monday, datetime.min.time()),
    )
    for header in EMAIL_CHECKLIST:
        cell_for(email_table, last_week_row, header).Value2 = True
    complete_links(email_table, last_week_row)
    cell_for(email_table, last_week_row, "Delivered").Value2 = 100

    set_common_row(sms_table, sms_row, QA_SMS, datetime.now())
    for header in SMS_CHECKLIST:
        cell_for(sms_table, sms_row, header).Value2 = False
    sms_stage = cell_for(sms_table, sms_row, "Current Stage")
    sms_stage.Calculate()
    sms_initial_stage = str(sms_stage.Value2 or "")
    cell_for(sms_table, sms_row, "Send SMS Options").Value2 = True
    sms_stage.Calculate()
    sms_second_stage = str(sms_stage.Value2 or "")
    for header in SMS_CHECKLIST:
        cell_for(sms_table, sms_row, header).Value2 = True
    complete_links(sms_table, sms_row)
    sms_stage.Calculate()
    sms_scheduled_stage = str(sms_stage.Value2 or "")

    run_macro(
        excel,
        workbook_name,
        "HandleCampaignChange",
        sms,
        cell_for(sms_table, sms_row, "Owner"),
    )
    sms_timestamp = cell_for(sms_table, sms_row, "Last Updated").Value2
    sms_user = str(
        cell_for(sms_table, sms_row, "Last Updated By").Value2 or ""
    ).strip()

    run_macro(excel, workbook_name, "RefreshDashboard")
    for sheet in workbook.Worksheets:
        sheet.Calculate()
    excel.Calculation = XL_CALCULATION_AUTOMATIC
    dashboard = workbook.Worksheets("Dashboard")
    dashboard_table = dashboard.ListObjects(DASHBOARD_TABLE)
    dashboard_campaigns = [
        str(value or "")
        for value in flatten(
            dashboard_table.ListColumns("Campaign").DataBodyRange.Value2
        )
    ]
    dashboard_channels = [
        str(value or "")
        for value in flatten(
            dashboard_table.ListColumns("Channel").DataBodyRange.Value2
        )
    ]

    calendar_name = date.today().strftime("%B Calendar")
    calendar_text = "\n".join(
        str(value or "")
        for value in flatten(
            workbook.Worksheets(calendar_name).Range("A6:G11").Value2
        )
    )

    expected_last_week = sum_delivered_for_week(email_table, previous_monday)
    expected_current_week = sum_delivered_for_week(email_table, current_monday)
    comparison = dashboard.ListObjects(COMPARISON_TABLE)
    actual_last_week = float(
        comparison.ListColumns("Delivered Emails").DataBodyRange.Cells(1, 1).Value2
        or 0
    )
    actual_current_week = float(
        comparison.ListColumns("Delivered Emails").DataBodyRange.Cells(2, 1).Value2
        or 0
    )

    validation = str(
        run_macro(excel, workbook_name, "ValidateWorkbookConfiguration")
    )
    workbook.Save()
    return {
        "email_stage_source_code": email_initial_stage == "Source Code",
        "email_stage_changes_dynamically": email_second_stage == "Creative Brief",
        "email_stage_scheduled": email_scheduled_stage == "Scheduled",
        "email_stage_sent": email_sent_stage == "Sent",
        "sms_stage_options": sms_initial_stage == "SMS Options",
        "sms_stage_changes_dynamically": sms_second_stage == "Send Test",
        "sms_stage_scheduled": sms_scheduled_stage == "Scheduled",
        "email_audit_fields": bool(email_timestamp) and bool(email_user),
        "sms_audit_fields": bool(sms_timestamp) and bool(sms_user),
        "dashboard_contains_email": QA_EMAIL in dashboard_campaigns,
        "dashboard_contains_sms": QA_SMS in dashboard_campaigns,
        "dashboard_channel_labels": "Email" in dashboard_channels
        and "SMS" in dashboard_channels,
        "calendar_contains_email": f"Email | {QA_EMAIL}" in calendar_text,
        "calendar_contains_sms": f"SMS | {QA_SMS}" in calendar_text,
        "last_week_delivered_total": abs(
            actual_last_week - expected_last_week
        )
        < 0.001,
        "current_week_delivered_total": abs(
            actual_current_week - expected_current_week
        )
        < 0.001,
        "vba_validation": validation == "OK",
    }


def persistence_checks(excel: Any, path: Path) -> dict[str, bool]:
    excel.AutomationSecurity = 1
    workbook = com_retry(
        lambda: excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False)
    )
    try:
        excel.AutomationSecurity = 1
        email_table = workbook.Worksheets(EMAIL_SHEET).ListObjects(EMAIL_TABLE)
        sms_table = workbook.Worksheets(SMS_SHEET).ListObjects(SMS_TABLE)
        email_found = email_table.ListColumns("Campaign Name").DataBodyRange.Find(
            QA_EMAIL
        )
        sms_found = sms_table.ListColumns("Campaign Name").DataBodyRange.Find(QA_SMS)
        validation = str(
            run_macro(
                excel,
                workbook.Name,
                "ValidateWorkbookConfiguration",
            )
        )
        return {
            "email_test_row_persisted": email_found is not None,
            "sms_test_row_persisted": sms_found is not None,
            "automatic_calculation_persisted": (
                excel.Calculation == XL_CALCULATION_AUTOMATIC
            ),
            "embedded_vba_persisted": validation == "OK",
        }
    finally:
        com_retry(lambda: workbook.Close(SaveChanges=False))


def open_excel(visible: bool) -> Any:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False
    excel.ScreenUpdating = visible
    excel.EnableEvents = False
    excel.AutomationSecurity = 3
    dummy = excel.Workbooks.Add()
    excel.Calculation = -4135
    dummy.Saved = True
    excel.AutomationSecurity = 1
    return excel


def package_uses_automatic_calculation(path: Path) -> bool:
    with ZipFile(path) as archive:
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
    match = re.search(r"<calcPr\b([^>]*)/?>", workbook_xml)
    if match is None:
        return True
    attributes = match.group(1)
    mode = re.search(r'calcMode="([^"]+)"', attributes)
    return mode is None or mode.group(1).lower() == "auto"


def run_qa(source: Path, visible: bool) -> dict[str, float]:
    working = source.with_name(f".{source.stem}.{uuid.uuid4().hex}.qa{source.suffix}")
    shutil.copy2(source, working)
    pythoncom.CoInitialize()
    excel = workbook = None
    checks: dict[str, bool] = {}
    timings: dict[str, float] = {}
    try:
        excel = open_excel(visible)
        started = time.perf_counter()
        workbook = com_retry(
            lambda: excel.Workbooks.Open(str(working), UpdateLinks=0, ReadOnly=False)
        )
        excel.AutomationSecurity = 1
        timings["open_seconds"] = time.perf_counter() - started
        print("  workbook_opened", flush=True)

        workbook_name = workbook.Name
        started = time.perf_counter()
        run_macro(excel, workbook_name, "ApplyAllConfigurations")
        print("  configuration_applied", flush=True)
        excel.Calculation = XL_CALCULATION_AUTOMATIC
        checks.update(structural_checks(excel, workbook))
        print("  structural_checks_complete", flush=True)
        excel.Calculation = -4135
        checks.update(behavior_checks(excel, workbook))
        print("  behavior_checks_complete", flush=True)
        timings["configuration_and_behavior_seconds"] = (
            time.perf_counter() - started
        )

        com_retry(lambda: workbook.Close(SaveChanges=False))
        workbook = None
        started = time.perf_counter()
        checks.update(persistence_checks(excel, working))
        checks["automatic_calculation_persisted"] = (
            package_uses_automatic_calculation(working)
        )
        timings["reopen_seconds"] = time.perf_counter() - started

        failed = [name for name, passed in checks.items() if not passed]
        for name, passed in checks.items():
            print(f"  {name}: {'OK' if passed else 'FAILED'}")
        for name, seconds in timings.items():
            print(f"  {name}: {seconds:.2f}s")
        if failed:
            raise RuntimeError("QA checks failed: " + ", ".join(failed))
        print(f"QA passed: {len(checks)} checks.")
        return timings
    finally:
        if workbook is not None:
            try:
                com_retry(lambda: workbook.Close(SaveChanges=False))
            except Exception:
                pass
        if excel is not None:
            try:
                com_retry(excel.Quit)
            except Exception:
                pass
        del workbook, excel
        gc.collect()
        pythoncom.CoUninitialize()
        working.unlink(missing_ok=True)


def apply_transactionally(source: Path, visible: bool) -> None:
    working = source.with_name(
        f".{source.stem}.{uuid.uuid4().hex}.apply{source.suffix}"
    )
    shutil.copy2(source, working)
    pythoncom.CoInitialize()
    excel = workbook = None
    try:
        excel = open_excel(visible)
        workbook = com_retry(
            lambda: excel.Workbooks.Open(str(working), UpdateLinks=0, ReadOnly=False)
        )
        excel.AutomationSecurity = 1
        run_macro(excel, workbook.Name, "ApplyAllConfigurations")
        validation = str(
            run_macro(excel, workbook.Name, "ValidateWorkbookConfiguration")
        )
        if validation != "OK":
            raise RuntimeError(validation)
        workbook.ForceFullCalculation = False
        for sheet in workbook.Worksheets:
            sheet.Calculate()
        excel.Calculation = XL_CALCULATION_AUTOMATIC
        workbook.Save()
        workbook.Close(SaveChanges=False)
        workbook = None
        excel.Quit()
        excel = None
        os.replace(working, source)
        print("Workbook configurations applied transactionally.")
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        del workbook, excel
        gc.collect()
        pythoncom.CoUninitialize()
        working.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Deployment QA and maintenance for the campaign tracker."
    )
    action = parser.add_mutually_exclusive_group()
    action.add_argument("--qa", action="store_true")
    action.add_argument("--apply", action="store_true")
    parser.add_argument("--visible", action="store_true")
    parser.add_argument(
        "--workbook",
        type=Path,
        default=Path(__file__).with_name(WORKBOOK_NAME),
    )
    args = parser.parse_args()
    path = args.workbook.resolve()
    if not path.exists():
        print(f"ERROR: Workbook not found: {path}", file=sys.stderr)
        return 1
    try:
        if args.apply:
            apply_transactionally(path, args.visible)
        else:
            run_qa(path, args.visible)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
