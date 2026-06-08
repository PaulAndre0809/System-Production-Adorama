"""Extensive QA harness for Email & SMS Campaign Tracker.xlsm.

This script copies the workbook to a temporary file, seeds 10 Email and
10 SMS test campaigns, exercises the embedded VBA and native formulas, and
then discards the temporary workbook. It is not required by workbook users.
"""

from __future__ import annotations

import argparse
import datetime as dt
import shutil
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client as win32


XL_CALC_AUTOMATIC = -4105
XL_CELL_TYPE_FORMULAS = -4123
XL_ERRORS = 16

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

EMAIL_CHECKLIST = [
    "Campaign Name and UTM Parameter (Source Code)",
    "Creative Brief, SL & PH",
    "SKUs",
    "In-Design",
    "Build, QA",
    "Route",
    "Approval",
    "Segments",
]

SMS_CHECKLIST = [
    "Send SMS Options",
    "Send Test",
    "Approval",
    "Segments",
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


def col_index(list_object, header_name: str) -> int:
    column = column_by_header(list_object, header_name)
    if column is None:
        raise AssertionError(f"Missing column {header_name!r} on {list_object.Name}")
    return int(column.Index)


def set_table_value(list_row, list_object, header_name: str, value) -> None:
    list_row.Range.Cells(1, col_index(list_object, header_name)).Value = value


def get_table_value(list_row, list_object, header_name: str):
    return list_row.Range.Cells(1, col_index(list_object, header_name)).Value


def flatten_values(value):
    if value is None:
        return []
    if not isinstance(value, tuple):
        return [value]
    result = []
    for item in value:
        result.extend(flatten_values(item))
    return result


def sheet_text(sheet) -> str:
    return "\n".join(str(value) for value in flatten_values(sheet.UsedRange.Value) if value is not None)


def excel_date(value) -> dt.date | None:
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)):
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(value))
    try:
        parsed = dt.datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        return parsed.date()
    except Exception:
        return None


def week_start_sunday(today: dt.date) -> dt.date:
    return today - dt.timedelta(days=(today.weekday() + 1) % 7)


def first_of_previous_month(today: dt.date) -> dt.date:
    first = today.replace(day=1)
    return (first - dt.timedelta(days=1)).replace(day=1)


def time_fraction(hour: int, minute: int = 0) -> float:
    return (hour * 60 + minute) / (24 * 60)


def excel_date_serial(value: dt.date) -> int:
    return (value - dt.date(1899, 12, 30)).days


def assert_zip_integrity(path: Path, checks: list[str]) -> None:
    with zipfile.ZipFile(path) as archive:
        bad_file = archive.testzip()
        if bad_file:
            raise AssertionError(f"Workbook zip integrity failed at {bad_file}")
        if "xl/vbaProject.bin" not in set(archive.namelist()):
            raise AssertionError("Workbook is missing embedded VBA project")
    checks.append("zip integrity and embedded VBA project")


def clear_table_filters(worksheet, table) -> None:
    try:
        if worksheet.FilterMode:
            worksheet.ShowAllData()
    except Exception:
        pass
    try:
        table.ShowAutoFilter = True
    except Exception:
        pass


def seed_campaign_rows(table, channel: str, today: dt.date) -> dict[str, list]:
    ws = week_start_sunday(today)
    previous_month = first_of_previous_month(today)
    now = dt.datetime.now().replace(microsecond=0)
    rows = []

    base = [
        (1, ws, "Promo", 0, True, True, "current week open"),
        (2, ws + dt.timedelta(days=1), "Services", 1250, True, True, "current week completed"),
        (3, ws + dt.timedelta(days=3), "Custom QA Type", 0, False, True, "custom campaign type"),
        (4, ws + dt.timedelta(days=4), None, 0, True, False, "blank campaign type"),
        (5, ws + dt.timedelta(days=8), "Loyalty & PLCC", 0, False, False, "next week open"),
        (6, ws + dt.timedelta(days=12), "Events", 0, True, True, "next week open"),
        (7, ws - dt.timedelta(days=2), "Newsletters", 700, True, True, "last week delivered"),
        (8, previous_month + dt.timedelta(days=5), "NPA", 0, False, False, "previous month open"),
        (9, previous_month + dt.timedelta(days=8), "Promo", 425, True, True, "previous month completed"),
        (10, ws + dt.timedelta(days=21), "Others", 0, False, True, "future outside dashboard"),
    ]

    for index, send_date, campaign_type, delivered, approval, segments, note in base:
        list_row = table.ListRows.Add()
        name = f"QA {channel} Campaign {index:02d}"
        set_table_value(list_row, table, "Send Date", dt.datetime.combine(send_date, dt.time()))
        set_table_value(list_row, table, "Send Time", time_fraction(8 + (index % 8), 30 if index % 2 else 0))
        set_table_value(list_row, table, "Campaign Name", name)
        set_table_value(list_row, table, "Campaign Type", campaign_type)
        set_table_value(list_row, table, "Owner", "QA Harness")
        set_table_value(list_row, table, "Jira Link", f"https://qa.example.com/jira/{channel.lower()}-{index:02d}")
        set_table_value(list_row, table, "ClickUp Link", f"https://qa.example.com/clickup/{channel.lower()}-{index:02d}")
        set_table_value(list_row, table, "Bluecore/Attentive Link", f"https://qa.example.com/build/{channel.lower()}-{index:02d}")
        set_table_value(list_row, table, "Est. Audience", 10000 + index * 1000)
        set_table_value(list_row, table, "Delivered", delivered)
        set_table_value(list_row, table, "Last Updated", now)
        set_table_value(list_row, table, "Last Updated By", "QA Harness")
        set_table_value(list_row, table, "Notes", f"QA seed row: {note}")

        if channel == "Email":
            set_table_value(list_row, table, "Campaign Name and UTM Parameter (Source Code)", index in (1, 2, 3, 4, 5, 6))
            set_table_value(list_row, table, "Creative Brief, SL & PH", index in (1, 2, 3, 6))
            set_table_value(list_row, table, "SKUs", index in (2, 4, 6))
            set_table_value(list_row, table, "In-Design", index in (3, 5))
            set_table_value(list_row, table, "Build, QA", index in (6,))
            set_table_value(list_row, table, "Route", index in (6,))
            set_table_value(list_row, table, "Approval", approval)
            set_table_value(list_row, table, "Segments", segments)
        else:
            set_table_value(list_row, table, "Send SMS Options", index in (1, 2, 3, 4, 5, 6))
            set_table_value(list_row, table, "Send Test", index in (2, 3, 6))
            set_table_value(list_row, table, "Approval", approval)
            set_table_value(list_row, table, "Segments", segments)

        rows.append({"name": name, "row": list_row, "date": send_date, "delivered": delivered})

    return {
        "all": rows,
        "dashboard_expected": [
            row for row in rows if ws <= row["date"] <= ws + dt.timedelta(days=13)
        ],
        "dashboard_excluded": [
            row for row in rows if not (ws <= row["date"] <= ws + dt.timedelta(days=13))
        ],
        "filter_visible": [
            row for row in rows if row["date"] >= today.replace(day=1) and row["delivered"] <= 0
        ],
    }


def table_rows_by_campaign(dashboard_table) -> dict[str, tuple]:
    values = dashboard_table.DataBodyRange.Value
    rows = values if isinstance(values, tuple) else ((values,),)
    result = {}
    for row in rows:
        if len(row) >= 4 and row[3] is not None and str(row[3]).strip():
            result[str(row[3])] = row
    return result


def assert_no_formula_errors(workbook, checks: list[str]) -> None:
    bad_cells = []
    for sheet_index in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets(sheet_index)
        try:
            errors = sheet.UsedRange.SpecialCells(XL_CELL_TYPE_FORMULAS, XL_ERRORS)
            bad_cells.append(f"{sheet.Name}!{errors.Address}")
        except pywintypes.com_error:
            pass
    if bad_cells:
        raise AssertionError("Formula errors found: " + ", ".join(bad_cells))

    bad_names = []
    for name in workbook.Names:
        try:
            refers_to = str(name.RefersTo)
            if "#REF!" in refers_to.upper():
                bad_names.append(str(name.Name))
        except Exception:
            pass
    if bad_names:
        raise AssertionError("Named ranges contain #REF!: " + ", ".join(bad_names))
    checks.append("no formula errors or broken named ranges")


def assert_campaign_type_sources(workbook, email_table, sms_table, checks: list[str]) -> None:
    dropdowns = workbook.Worksheets("Dropdowns")
    actual = [dropdowns.Cells(row, 1).Value for row in range(2, 10)]
    if actual != EXPECTED_CAMPAIGN_TYPES:
        raise AssertionError(f"Campaign Type dropdown source mismatch: {actual}")
    for table in (email_table, sms_table):
        col = column_by_header(table, "Campaign Type")
        validation = col.DataBodyRange.Validation
        if validation.Type != 3 or validation.Formula1 != "=Dropdowns!$A$2:$A$9":
            raise AssertionError(f"Campaign Type validation is wrong on {table.Name}")
        if validation.ShowError:
            raise AssertionError(f"Campaign Type custom values are blocked on {table.Name}")
    checks.append("campaign type dropdowns, blank option, and custom value behavior")


def assert_checklist_columns(table, checklist_headers: list[str], checks: list[str]) -> None:
    for header in checklist_headers:
        col = column_by_header(table, header)
        if col is None or col.DataBodyRange is None:
            raise AssertionError(f"Missing checklist column {header!r} on {table.Name}")
        for row_index in range(1, min(12, col.DataBodyRange.Rows.Count) + 1):
            value = col.DataBodyRange.Cells(row_index, 1).Value
            if value not in (True, False):
                raise AssertionError(f"Checklist value is not boolean: {table.Name}[{header}] row {row_index}")
        try:
            validation_type = col.DataBodyRange.Cells(1, 1).Validation.Type
            if validation_type == 3:
                raise AssertionError(f"Checklist column still has dropdown validation: {table.Name}[{header}]")
        except Exception:
            pass
    checks.append(f"{table.Name} checklist columns are boolean-only")


def assert_stage_values(table, seeded_rows: list[dict], channel: str, checks: list[str]) -> None:
    for row_info in seeded_rows:
        value = str(get_table_value(row_info["row"], table, "Current Stage") or "")
        if not value:
            raise AssertionError(f"Current Stage is blank for {row_info['name']}")
        if row_info["name"].endswith("01") and "Checked:" not in value:
            raise AssertionError(f"Current Stage did not list checked items for {row_info['name']}: {value}")
        if channel == "Email" and row_info["name"].endswith("01") and "Campaign Name and UTM Parameter" not in value:
            raise AssertionError(f"Email Current Stage missing checked source-code step: {value}")
        if channel == "SMS" and row_info["name"].endswith("01") and "Send SMS Options" not in value:
            raise AssertionError(f"SMS Current Stage missing checked SMS step: {value}")
    checks.append(f"{channel} Current Stage formulas update from checked columns")


def assert_dashboard_rows(dashboard_table, email_seed, sms_seed, checks: list[str]) -> None:
    rows_by_campaign = table_rows_by_campaign(dashboard_table)

    expected = email_seed["dashboard_expected"] + sms_seed["dashboard_expected"]
    excluded = email_seed["dashboard_excluded"] + sms_seed["dashboard_excluded"]
    for row_info in expected:
        if row_info["name"] not in rows_by_campaign:
            raise AssertionError(f"Dashboard missing expected campaign: {row_info['name']}")
        row = rows_by_campaign[row_info["name"]]
        approval = row[7]
        segments = row[8]
        if approval not in ("Done", "Not Yet"):
            raise AssertionError(f"Dashboard approval status is not friendly text for {row_info['name']}: {approval}")
        if segments not in ("Provided", "Pending"):
            raise AssertionError(f"Dashboard segments status is not friendly text for {row_info['name']}: {segments}")
    for row_info in excluded:
        if row_info["name"] in rows_by_campaign:
            raise AssertionError(f"Dashboard included out-of-window campaign: {row_info['name']}")
    checks.append("Dashboard current-week through next-week feed and friendly statuses")


def assert_calendar_rows(workbook, email_seed, sms_seed, today: dt.date, checks: list[str]) -> None:
    current_sheet = workbook.Worksheets(today.strftime("%B") + " Calendar")
    current_text = sheet_text(current_sheet)
    for row_info in email_seed["dashboard_expected"][:2] + sms_seed["dashboard_expected"][:2]:
        if row_info["date"].month == today.month and row_info["name"] not in current_text:
            raise AssertionError(f"Current month calendar missing {row_info['name']}")

    previous_month = first_of_previous_month(today)
    previous_sheet = workbook.Worksheets(previous_month.strftime("%B") + " Calendar")
    previous_text = sheet_text(previous_sheet)
    previous_names = [
        row["name"]
        for row in email_seed["all"] + sms_seed["all"]
        if row["date"].month == previous_month.month
    ]
    for name in previous_names[:2]:
        if name not in previous_text:
            raise AssertionError(f"Previous month calendar missing {name}")

    expected_current_color = 5287936  # RGB(0, 176, 80)
    if int(current_sheet.Tab.Color) != expected_current_color:
        raise AssertionError("Current month calendar tab is not green")
    checks.append("calendar sheets include Email/SMS rows and highlight current month")


def assert_filters(table, seed, today: dt.date, checks: list[str]) -> None:
    worksheet = table.Parent
    clear_table_filters(worksheet, table)

    send_date_field = col_index(table, "Send Date")
    delivered_field = col_index(table, "Delivered")
    month_start_serial = excel_date_serial(today.replace(day=1))
    table.Range.AutoFilter(Field=send_date_field, Criteria1=f">={month_start_serial}")
    table.Range.AutoFilter(Field=delivered_field, Criteria1="<=0")

    hidden_by_name = {}
    for row_info in seed["all"]:
        hidden_by_name[row_info["name"]] = bool(row_info["row"].Range.EntireRow.Hidden)

    for row_info in seed["filter_visible"]:
        if hidden_by_name.get(row_info["name"]):
            raise AssertionError(f"Filter hid an expected visible row: {row_info['name']}")

    for row_info in seed["all"]:
        should_hide = row_info not in seed["filter_visible"]
        if should_hide and not hidden_by_name.get(row_info["name"]):
            raise AssertionError(f"Filter did not hide completed/previous-month row: {row_info['name']}")

    clear_table_filters(worksheet, table)
    checks.append(f"{table.Name} filters can hide completed and previous-month campaigns")


def assert_delivered_comparison(workbook, email_table, today: dt.date, checks: list[str]) -> None:
    dashboard = workbook.Worksheets("Dashboard")
    week_start = week_start_sunday(today)
    last_start = week_start - dt.timedelta(days=7)
    last_end = week_start - dt.timedelta(days=1)
    current_end = week_start + dt.timedelta(days=6)

    date_col = col_index(email_table, "Send Date")
    delivered_col = col_index(email_table, "Delivered")
    expected_last = 0
    expected_current = 0
    values = email_table.DataBodyRange.Value2
    for row in values:
        row_date = excel_date(row[date_col - 1])
        delivered = row[delivered_col - 1] or 0
        if row_date is None:
            continue
        if last_start <= row_date <= last_end:
            expected_last += float(delivered)
        if week_start <= row_date <= current_end:
            expected_current += float(delivered)

    actual_last = float(dashboard.Range("Q4").Value or 0)
    actual_current = float(dashboard.Range("Q5").Value or 0)
    if round(actual_last, 4) != round(expected_last, 4):
        raise AssertionError(f"Delivered comparison last week mismatch: {actual_last} vs {expected_last}")
    if round(actual_current, 4) != round(expected_current, 4):
        raise AssertionError(f"Delivered comparison current week mismatch: {actual_current} vs {expected_current}")
    if excel_date(dashboard.Range("O5").Value) != week_start or excel_date(dashboard.Range("P5").Value) != current_end:
        raise AssertionError("Delivered comparison is not using Sunday-Saturday current week")
    if not any(chart.Name == "DeliveredEmailComparisonChart" for chart in dashboard.ChartObjects()):
        raise AssertionError("Delivered email comparison chart is missing")
    checks.append("Delivered comparison chart/table uses Sunday-Saturday weeks")


def assert_dashboard_audit_header(dashboard, checks: list[str]) -> None:
    if dashboard.Range("A3").Value != "Last Refresh":
        raise AssertionError("Dashboard Last Refresh label missing")
    if dashboard.Range("C3").Value != "Last Edited By":
        raise AssertionError("Dashboard Last Edited By label missing")
    if str(dashboard.Range("B3").Formula) != '=TEXT(NOW(),"m/d/yyyy h:mm AM/PM")':
        raise AssertionError("Dashboard Last Refresh formula is wrong")
    if dashboard.Range("E3").Value is not None or dashboard.Range("F3").Value is not None:
        raise AssertionError("Dashboard still contains leftover Last Edited cells")
    checks.append("Dashboard audit header has Last Refresh and Last Edited By only")


def assert_formats_and_ui(workbook, email_table, sms_table, checks: list[str]) -> None:
    for table in (email_table, sms_table):
        for header, expected in (
            ("Send Date", "mm/dd/yyyy"),
            ("Send Time", "am/pm"),
            ("Last Updated", "am/pm"),
            ("Last Updated By", "@"),
        ):
            col = column_by_header(table, header)
            actual = str(col.DataBodyRange.NumberFormat).lower()
            if expected == "@":
                if actual != "@":
                    raise AssertionError(f"{table.Name}[{header}] should be text format")
            elif expected not in actual:
                raise AssertionError(f"{table.Name}[{header}] format is wrong: {actual}")
        if not table.ShowAutoFilter:
            raise AssertionError(f"{table.Name} filter dropdowns are disabled")

    for sheet_index in range(1, workbook.Worksheets.Count + 1):
        sheet = workbook.Worksheets(sheet_index)
        sheet.Activate()
        if workbook.Application.ActiveWindow.FreezePanes:
            raise AssertionError(f"Freeze panes are enabled on {sheet.Name}")
        if workbook.Application.ActiveWindow.SplitRow or workbook.Application.ActiveWindow.SplitColumn:
            raise AssertionError(f"Split panes are enabled on {sheet.Name}")
    checks.append("date/time formats, table filters, and unfrozen views")


def assert_links_and_hidden_helpers(workbook, dashboard, checks: list[str]) -> None:
    links = retry("external links", lambda: workbook.LinkSources(1))
    if links is not None:
        raise AssertionError(f"External workbook links detected: {links}")
    if not dashboard.Columns("AA:AL").Hidden:
        raise AssertionError("Dashboard helper columns AA:AL should be hidden")
    if not str(dashboard.Range("AA11").Formula2).startswith("=LET("):
        raise AssertionError("Dashboard native helper formula is missing")
    checks.append("hidden Dashboard helpers and no external workbook links")


def validate_extensively(path: Path) -> tuple[list[str], dict[str, float]]:
    checks: list[str] = []
    timings: dict[str, float] = {}
    assert_zip_integrity(path, checks)

    temp_dir = Path(tempfile.mkdtemp(prefix="campaign_tracker_extensive_qa_"))
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
            excel.Calculation = XL_CALC_AUTOMATIC
        except Exception:
            pass

        start = time.perf_counter()
        workbook = retry(
            "open workbook",
            lambda: excel.Workbooks.Open(str(qa_path), UpdateLinks=0, ReadOnly=False, AddToMru=False),
        )
        timings["open_seconds"] = time.perf_counter() - start

        start = time.perf_counter()
        validation = retry(
            "embedded validation",
            lambda: excel.Run(macro_name(workbook, "ValidateWorkbookConfiguration")),
        )
        timings["validate_macro_seconds"] = time.perf_counter() - start
        if validation != "OK":
            raise AssertionError(f"Embedded validation failed: {validation}")
        checks.append("embedded validation macro")

        start = time.perf_counter()
        retry("ApplyAllConfigurations", lambda: excel.Run(macro_name(workbook, "ApplyAllConfigurations")))
        timings["apply_all_configurations_seconds"] = time.perf_counter() - start
        checks.append("ApplyAllConfigurations macro")

        dashboard = workbook.Worksheets("Dashboard")
        email_ws = workbook.Worksheets("Email Campaigns")
        sms_ws = workbook.Worksheets("SMS Campaigns")
        email_table = email_ws.ListObjects("EmailCampaignsTable")
        sms_table = sms_ws.ListObjects("SMSCampaignsTable")
        dashboard_table = dashboard.ListObjects("DashboardWorkTable")
        today = dt.date.today()

        clear_table_filters(email_ws, email_table)
        clear_table_filters(sms_ws, sms_table)

        email_seed = seed_campaign_rows(email_table, "Email", today)
        sms_seed = seed_campaign_rows(sms_table, "SMS", today)
        checks.append("seeded 10 Email and 10 SMS QA rows in temporary copy")

        start = time.perf_counter()
        retry("RefreshDashboard", lambda: excel.Run(macro_name(workbook, "RefreshDashboard")))
        timings["refresh_dashboard_seconds"] = time.perf_counter() - start

        start = time.perf_counter()
        retry("RefreshNativeOutputs", lambda: excel.Run(macro_name(workbook, "RefreshNativeOutputs")))
        timings["refresh_native_outputs_seconds"] = time.perf_counter() - start

        start = time.perf_counter()
        retry("CalculateFull", lambda: excel.CalculateFull())
        timings["calculate_full_seconds"] = time.perf_counter() - start

        validation = retry(
            "post-seed embedded validation",
            lambda: excel.Run(macro_name(workbook, "ValidateWorkbookConfiguration")),
        )
        if validation != "OK":
            raise AssertionError(f"Post-seed embedded validation failed: {validation}")
        checks.append("post-seed embedded validation macro")

        assert_dashboard_audit_header(dashboard, checks)
        assert_campaign_type_sources(workbook, email_table, sms_table, checks)
        assert_checklist_columns(email_table, EMAIL_CHECKLIST, checks)
        assert_checklist_columns(sms_table, SMS_CHECKLIST, checks)
        assert_stage_values(email_table, email_seed["all"], "Email", checks)
        assert_stage_values(sms_table, sms_seed["all"], "SMS", checks)
        assert_dashboard_rows(dashboard_table, email_seed, sms_seed, checks)
        assert_calendar_rows(workbook, email_seed, sms_seed, today, checks)
        assert_filters(email_table, email_seed, today, checks)
        assert_filters(sms_table, sms_seed, today, checks)
        assert_delivered_comparison(workbook, email_table, today, checks)
        assert_formats_and_ui(workbook, email_table, sms_table, checks)
        assert_links_and_hidden_helpers(workbook, dashboard, checks)
        assert_no_formula_errors(workbook, checks)

        # Exercise the explicit checkbox toggle macro on one seeded Email and SMS row.
        for table, seed, header in (
            (email_table, email_seed, "Approval"),
            (sms_table, sms_seed, "Approval"),
        ):
            target = seed["all"][0]["row"].Range.Cells(1, col_index(table, header))
            old_value = bool(target.Value)
            result = retry(
                f"ToggleInventoryChecklist {table.Name}",
                lambda target=target: excel.Run(macro_name(workbook, "ToggleInventoryChecklist"), target),
            )
            if result is not True or bool(target.Value) == old_value:
                raise AssertionError(f"ToggleInventoryChecklist failed on {table.Name}")
        checks.append("checkbox toggle macro")

        workbook.Close(SaveChanges=False)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        shutil.rmtree(temp_dir, ignore_errors=True)

    return checks, timings


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

    checks, timings = validate_extensively(workbook_path)
    print("EXTENSIVE QA PASSED")
    for check in checks:
        print(f"- {check}")
    print("Timings:")
    for key, value in timings.items():
        print(f"- {key}: {value:.2f}s")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
