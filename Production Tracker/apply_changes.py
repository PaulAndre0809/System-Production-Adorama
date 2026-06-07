from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client


MODULE = "modEmailProductionTracker"
PREFERRED_WORKBOOK = "Email & SMS Campaign Tracker.xlsm"
LEGACY_WORKBOOK = "Email_Production_Inventory_Tracker_UI.xlsm"
RPC_E_CALL_REJECTED = -2147418111


def com_retry(operation, attempts: int = 40, delay: float = 0.25):
    for attempt in range(attempts):
        try:
            return operation()
        except pywintypes.com_error as exc:
            if exc.hresult != RPC_E_CALL_REJECTED or attempt == attempts - 1:
                raise
            time.sleep(delay)
    raise RuntimeError("Excel COM operation did not complete.")


def run_macro(excel, workbook_name: str, procedure: str, *args):
    macro_name = f"'{workbook_name}'!{MODULE}.{procedure}"
    return com_retry(lambda: excel.Run(macro_name, *args))


def apply_changes(visible: bool) -> None:
    workbook_path = Path(__file__).with_name(PREFERRED_WORKBOOK).resolve()
    if not workbook_path.exists():
        workbook_path = Path(__file__).with_name(LEGACY_WORKBOOK).resolve()
        if not workbook_path.exists():
            raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    pythoncom.CoInitialize()
    excel = None
    workbook = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = visible
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1

        workbook = com_retry(
            lambda: excel.Workbooks.Open(str(workbook_path))
        )
        workbook_name = com_retry(lambda: workbook.Name)
        run_macro(excel, workbook_name, "ApplyAllConfigurations")

        validation = str(
            run_macro(
                excel,
                workbook_name,
                "ValidateWorkbookConfiguration",
            )
        )
        if validation != "OK":
            raise RuntimeError(validation)

        com_retry(workbook.Save)
        print("Workbook configurations applied successfully.")
        print(f"Validation: {validation}")
    finally:
        if workbook is not None:
            com_retry(lambda: workbook.Close(SaveChanges=False))
        if excel is not None:
            com_retry(excel.Quit)
        pythoncom.CoUninitialize()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Apply and validate Production Tracker workbook settings."
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show Excel while the automation runs.",
    )
    args = parser.parse_args()

    try:
        apply_changes(args.visible)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
