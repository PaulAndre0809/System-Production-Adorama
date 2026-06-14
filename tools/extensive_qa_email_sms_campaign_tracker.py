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
XL_CALC_MANUAL = -4135
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


def set_table_format(list_row, list_object, header_name: str, number_format: str) -> None:
    list_row.Range.Cells(
        1, col_index(list_object, header_name)
    ).NumberFormat = number_format


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
        if index == 3:
            send_time = "STO"
        elif index == 4:
            send_time = "Local Timezone"
        else:
            send_time = time_fraction(
                8 + (index % 8), 30 if index % 2 else 0
            )
        set_table_value(list_row, table, "Send Time", send_time)
        set_table_value(list_row, table, "Campaign Name", name)
        set_table_value(list_row, table, "Campaign Type", campaign_type)
        set_table_value(list_row, table, "Owner", "QA Harness")
        if channel == "Email":
            set_table_value(list_row, table, "Jira Link", f"https://qa.example.com/jira/{channel.lower()}-{index:02d}")
            set_table_value(list_row, table, "ClickUp Link", f"https://qa.example.com/clickup/{channel.lower()}-{index:02d}")
        else:
            set_table_value(list_row, table, "Proof of Schedule", f"https://qa.example.com/proof/{channel.lower()}-{index:02d}")
        set_table_value(list_row, table, "Bluecore/Attentive Link", f"https://qa.example.com/build/{channel.lower()}-{index:02d}")
        set_table_value(list_row, table, "Est. Audience", 10000 + index * 1000)
        set_table_value(list_row, table, "Delivered", delivered)
        set_table_value(list_row, table, "Last Updated", now)
        set_table_value(list_row, table, "Last Updated By", "QA Harness")
        cancelled = index == 4
        set_table_value(
            list_row,
            table,
            "Notes",
            "  CANCELLED  " if cancelled else f"QA seed row: {note}",
        )
        set_table_format(
            list_row, table, "Send Date", "dddd, mmmm d, yyyy"
        )
        set_table_format(list_row, table, "Send Time", "h:mm AM/PM")
        set_table_format(
            list_row, table, "Last Updated", "MM/DD/YYYY h:mm AM/PM"
        )
        set_table_format(list_row, table, "Last Updated By", "@")

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

        rows.append(
            {
                "name": name,
                "row": list_row,
                "date": send_date,
                "delivered": delivered,
                "cancelled": cancelled,
            }
        )

    return {
        "all": rows,
        "dashboard_expected": [
            row
            for row in rows
            if (
                ws <= row["date"] <= ws + dt.timedelta(days=13)
                and not row["cancelled"]
            )
        ],
        "dashboard_excluded": [
            row
            for row in rows
            if (
                not (ws <= row["date"] <= ws + dt.timedelta(days=13))
                or row["cancelled"]
            )
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
        headers = EMAIL_CHECKLIST if channel == "Email" else SMS_CHECKLIST
        checked_headers = [
            header
            for header in headers
            if bool(get_table_value(row_info["row"], table, header))
        ]

        if not value:
            raise AssertionError(f"Current Stage is blank for {row_info['name']}")

        if checked_headers:
            if not value.startswith("Checked: "):
                raise AssertionError(
                    f"Current Stage did not list checked items for {row_info['name']}: {value}"
                )
            for header in checked_headers:
                if header not in value:
                    raise AssertionError(
                        f"Current Stage omitted {header!r} for {row_info['name']}: {value}"
                    )
        else:
            if value != "No checklist items checked":
                raise AssertionError(
                    f"Current Stage should show no checked items for {row_info['name']}: {value}"
                )
    checks.append(f"{channel} Current Stage lists every checked workflow column")


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


def source_kpi_counts(table, today: dt.date) -> dict[str, int]:
    """Calculate expected Dashboard KPI values from the source table."""
    rows = table.DataBodyRange.Value2
    rows = rows if isinstance(rows, tuple) else ((rows,),)
    indices = {
        str(table.ListColumns(index).Name): index - 1
        for index in range(1, table.ListColumns.Count + 1)
    }
    counts = {"active": 0, "today": 0, "approval": 0, "sent": 0}

    for row in rows:
        campaign = row[indices["Campaign Name"]]
        if campaign is None or not str(campaign).strip():
            continue
        notes = str(row[indices["Notes"]] or "").strip().lower()
        stage = str(row[indices["Current Stage"]] or "").strip().lower()
        cancelled = (
            notes in {"cancelled", "canceled"}
            or stage in {"cancelled", "canceled"}
        )
        delivered_value = row[indices["Delivered"]]
        try:
            delivered = float(delivered_value or 0)
        except (TypeError, ValueError):
            delivered = 0

        if delivered > 0 and not cancelled:
            counts["sent"] += 1
        elif not cancelled:
            counts["active"] += 1
        if not cancelled and excel_date(row[indices["Send Date"]]) == today:
            counts["today"] += 1
        if not cancelled and not bool(row[indices["Approval"]]):
            counts["approval"] += 1
    return counts


def assert_dashboard_kpis(
    dashboard,
    email_table,
    sms_table,
    today: dt.date,
    checks: list[str],
) -> None:
    """Verify all six KPI values and reject implicit-intersection formulas."""
    for address in ("C5", "E5", "G5", "I5", "K5"):
        formula = str(dashboard.Range(address).Formula2)
        if "@Delivered" in formula or "--@Email" in formula or "--@SMS" in formula:
            raise AssertionError(f"Dashboard KPI {address} uses implicit intersection")

    email = source_kpi_counts(email_table, today)
    sms = source_kpi_counts(sms_table, today)
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
    checks.append("Dashboard KPI formulas and values match all source rows")


def assert_retired_features(workbook, dashboard, checks: list[str]) -> None:
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
        if "deliveredcomparison" in str(dashboard.ListObjects(index).Name).lower()
    ]
    if comparison_tables:
        raise AssertionError(f"Retired comparison tables still exist: {comparison_tables}")

    comparison_charts = [
        dashboard.ChartObjects(index).Name
        for index in range(1, dashboard.ChartObjects().Count + 1)
        if "delivered" in str(dashboard.ChartObjects(index).Name).lower()
        and "comparison" in str(dashboard.ChartObjects(index).Name).lower()
    ]
    if comparison_charts:
        raise AssertionError(f"Retired comparison charts still exist: {comparison_charts}")

    if any(
        value is not None and str(value).strip()
        for value in flatten_values(dashboard.Range("N2:S44").Value)
    ):
        raise AssertionError("Retired Dashboard comparison cells are not empty")
    if any(
        value is not None and str(value).strip()
        for value in flatten_values(dashboard.Range("A7:L7").Value)
    ) or dashboard.Range("A7:L7").Hyperlinks.Count:
        raise AssertionError("Retired Dashboard calendar navigation is not empty")

    notes = workbook.Worksheets("Notes - Instructions")
    notes_text = sheet_text(notes)
    if notes.Range("A1").Value != "Detailed Notes and Instructions":
        raise AssertionError("Notes - Instructions title is outdated")
    if "intentionally removed" not in notes_text:
        raise AssertionError("Notes do not document the retired features")
    if "Wednesday, June 10, 2026" not in notes_text:
        raise AssertionError("Notes do not document the Send Date format")
    if "STO or Local Timezone" not in notes_text:
        raise AssertionError("Notes do not document allowed Send Time text")
    if "Timed Link Labels" not in notes_text:
        raise AssertionError("Notes do not document timed link labels")
    if "Cancelled Campaigns" not in notes_text:
        raise AssertionError("Notes do not document cancellation filtering")
    checks.append("Calendar sheets and weekly Dashboard comparisons remain retired")


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
        row_count = table.DataBodyRange.Rows.Count
        sample_rows = sorted(
            {1, 2, 3}
            | set(range(max(1, row_count - 11), row_count + 1))
        )
        for header, expected in (
            ("Send Date", "dddd, mmmm d, yyyy"),
            ("Send Time", "h:mm am/pm"),
            ("Last Updated", "am/pm"),
            ("Last Updated By", "@"),
        ):
            col = column_by_header(table, header)
            for row_index in sample_rows:
                actual = str(
                    col.DataBodyRange.Cells(row_index, 1).NumberFormat
                ).lower()
                if expected == "@":
                    if actual != "@":
                        raise AssertionError(
                            f"{table.Name}[{header}] row {row_index} should be text format"
                        )
                elif header in {"Send Date", "Send Time"} and expected != actual:
                    raise AssertionError(
                        f"{table.Name}[{header}] row {row_index} format is wrong: {actual}"
                    )
                elif header not in {"Send Date", "Send Time"} and expected not in actual:
                    raise AssertionError(
                        f"{table.Name}[{header}] row {row_index} format is wrong: {actual}"
                    )
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
    dashboard_formula = str(dashboard.Range("AA11").Formula2)
    if not dashboard_formula.startswith("=LET("):
        raise AssertionError("Dashboard native helper formula is missing")
    lowered_formula = dashboard_formula.lower()
    if "cancelled" not in lowered_formula or "[current stage]" not in lowered_formula:
        raise AssertionError("Dashboard native helper does not exclude cancelled rows")
    checks.append("hidden Dashboard helpers and no external workbook links")


def assert_timed_hyperlinks(
    workbook,
    email_table,
    sms_table,
    email_seed,
    sms_seed,
    checks: list[str],
) -> None:
    """Verify both sides of the exact seven-day link-label boundary."""
    retry(
        "ApplyTimedCampaignLinks",
        lambda: workbook.Application.Run(
            macro_name(workbook, "ApplyTimedCampaignLinks")
        ),
    )

    now = dt.datetime.now().replace(microsecond=0)
    before_maturity = now - dt.timedelta(days=7) + dt.timedelta(minutes=10)
    after_maturity = now - dt.timedelta(days=7) - dt.timedelta(minutes=10)

    for table, seed, configurations in (
        (
            email_table,
            email_seed,
            (
                ("Jira Link", "JIRA"),
                ("ClickUp Link", "ClickUp"),
                ("Bluecore/Attentive Link", "Bluecore/Attentive"),
            ),
        ),
        (
            sms_table,
            sms_seed,
            (
                ("Proof of Schedule", "Proof of Schedule"),
                ("Bluecore/Attentive Link", "Bluecore/Attentive"),
            ),
        ),
    ):
        test_row = seed["all"][9]["row"]
        set_table_value(
            test_row,
            table,
            "Send Date",
            dt.datetime.combine(before_maturity.date(), dt.time()),
        )
        set_table_value(
            test_row,
            table,
            "Send Time",
            time_fraction(before_maturity.hour, before_maturity.minute),
        )

        for header, display_name in configurations:
            cell = test_row.Range.Cells(1, col_index(table, header))
            address = str(cell.Value2 or "")
            if not address.lower().startswith("http"):
                formula = str(cell.Formula2 or "")
                match = formula.split('"', 2)
                address = match[1].replace('""', '"') if len(match) >= 3 else ""
            cell.Calculate()
            formula = str(cell.Formula2 or "")
            if (
                "HYPERLINK(" not in formula.upper()
                or "NOW()" not in formula.upper()
                or "+7" not in formula
                or address not in formula
            ):
                raise AssertionError(
                    f"Timed formula is incomplete on {table.Name}[{header}]"
                )
            if str(cell.Value2) != address:
                raise AssertionError(
                    f"{table.Name}[{header}] matured before seven full days"
                )

        set_table_value(
            test_row,
            table,
            "Send Date",
            dt.datetime.combine(after_maturity.date(), dt.time()),
        )
        set_table_value(
            test_row,
            table,
            "Send Time",
            time_fraction(after_maturity.hour, after_maturity.minute),
        )
        for header, display_name in configurations:
            cell = test_row.Range.Cells(1, col_index(table, header))
            cell.Calculate()
            if str(cell.Value2) != display_name:
                raise AssertionError(
                    f"{table.Name}[{header}] did not mature to {display_name!r}"
                )

        set_table_value(
            test_row,
            table,
            "Send Date",
            dt.datetime.combine(
                dt.date.today() - dt.timedelta(days=8),
                dt.time(),
            ),
        )
        set_table_value(test_row, table, "Send Time", "STO")
        for header, display_name in configurations:
            cell = test_row.Range.Cells(1, col_index(table, header))
            cell.Calculate()
            if str(cell.Value2) != display_name:
                raise AssertionError(
                    f"{table.Name}[{header}] text-time fallback is incorrect"
                )

    checks.append(
        "all supported hyperlinks switch at the seven-day timestamp, including text-time fallback"
    )


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
        excel.Calculation = XL_CALC_MANUAL

        start = time.perf_counter()
        validation = retry(
            "embedded validation",
            lambda: excel.Run(macro_name(workbook, "ValidateWorkbookConfiguration")),
        )
        timings["validate_macro_seconds"] = time.perf_counter() - start
        if validation != "OK":
            raise AssertionError(f"Embedded validation failed: {validation}")
        checks.append("embedded validation macro")

        retry(
            "legacy UpdateCalendarTabs wrapper",
            lambda: excel.Run(macro_name(workbook, "UpdateCalendarTabs")),
        )

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

        retry(
            "ApplyTimedCampaignLinks",
            lambda: excel.Run(macro_name(workbook, "ApplyTimedCampaignLinks")),
        )

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
        assert_dashboard_kpis(
            dashboard,
            email_table,
            sms_table,
            today,
            checks,
        )
        assert_retired_features(workbook, dashboard, checks)
        assert_filters(email_table, email_seed, today, checks)
        assert_filters(sms_table, sms_seed, today, checks)
        assert_formats_and_ui(workbook, email_table, sms_table, checks)
        assert_links_and_hidden_helpers(workbook, dashboard, checks)
        assert_timed_hyperlinks(
            workbook,
            email_table,
            sms_table,
            email_seed,
            sms_seed,
            checks,
        )
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
