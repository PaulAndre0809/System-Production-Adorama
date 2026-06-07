from __future__ import annotations

import shutil
import sys
import tempfile
import uuid
from datetime import datetime
from pathlib import Path

import pythoncom
import win32com.client

from apply_changes import MODULE, WORKBOOK, com_retry, run_macro


QA_CAMPAIGN = "QA TEMP CAMPAIGN"


def find_test_row(table) -> int:
    campaign_column = table.ListColumns("Campaign Name").DataBodyRange
    for index in range(table.ListRows.Count, 0, -1):
        value = campaign_column.Cells(index, 1).Value2
        if value is None or not str(value).strip():
            return index
    table.ListRows.Add()
    return table.ListRows.Count


def cell_for(table, row_index: int, header: str):
    column_index = table.ListColumns(header).Index
    return table.DataBodyRange.Rows(row_index).Cells(1, column_index)


def run_smoke_test() -> None:
    source = Path(__file__).with_name(WORKBOOK).resolve()
    working = Path(tempfile.gettempdir()) / (
        f"ProductionTracker_QA_{uuid.uuid4().hex}.xlsm"
    )
    shutil.copy2(source, working)

    pythoncom.CoInitialize()
    excel = None
    workbook = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1

        workbook = com_retry(lambda: excel.Workbooks.Open(str(working)))
        workbook_name = com_retry(lambda: workbook.Name)
        inventory = workbook.Worksheets("Production Inventory")
        table = inventory.ListObjects("ProductionInventoryTable")
        row_index = find_test_row(table)

        send_date = cell_for(table, row_index, "Send Date")
        campaign = cell_for(table, row_index, "Campaign Name")
        owner = cell_for(table, row_index, "Owner")
        checklist = cell_for(table, row_index, "SKUs")
        updated = cell_for(table, row_index, "Last Updated")
        updated_by = cell_for(table, row_index, "Last Updated By")
        stage = cell_for(table, row_index, "Current Stage")

        send_date.Value = datetime.now()
        campaign.Value2 = QA_CAMPAIGN
        owner.Value2 = "QA"

        run_macro(
            excel,
            workbook_name,
            "HandleInventoryChange",
            campaign,
        )
        before_toggle = str(checklist.Value2 or "")
        toggled = bool(
            run_macro(
                excel,
                workbook_name,
                "ToggleInventoryChecklist",
                checklist,
            )
        )
        after_toggle = str(checklist.Value2 or "")
        timestamp_before_refresh = float(updated.Value2 or 0)

        run_macro(excel, workbook_name, "RefreshProductionStatus")
        timestamp_after_refresh = float(updated.Value2 or 0)
        inventory.Calculate()

        validation = str(
            run_macro(
                excel,
                workbook_name,
                "ValidateWorkbookConfiguration",
            )
        )
        dashboard = workbook.Worksheets("Dashboard")
        found = dashboard.Columns("C").Find(QA_CAMPAIGN)

        checks = {
            "validation": validation == "OK",
            "timestamp_created": timestamp_before_refresh > 0,
            "timestamp_preserved": abs(
                timestamp_after_refresh - timestamp_before_refresh
            )
            < 0.0000001,
            "updated_by": bool(str(updated_by.Value2 or "").strip()),
            "checklist_toggled": toggled and after_toggle != before_toggle,
            "stage": str(stage.Value2 or "") == "Source Code",
            "dashboard": found is not None,
        }

        failed = [name for name, passed in checks.items() if not passed]
        if failed:
            raise RuntimeError("Smoke checks failed: " + ", ".join(failed))

        print("Workbook smoke test passed.")
        for name in checks:
            print(f"  {name}: OK")
    finally:
        if workbook is not None:
            com_retry(lambda: workbook.Close(SaveChanges=False))
        if excel is not None:
            com_retry(excel.Quit)
        pythoncom.CoUninitialize()
        working.unlink(missing_ok=True)


def main() -> int:
    try:
        run_smoke_test()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
