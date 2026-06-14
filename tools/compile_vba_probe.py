"""Compile an Excel workbook's VBA project and detect modal compiler errors."""

from __future__ import annotations

import sys
import threading
import time
from pathlib import Path

import pythoncom
import win32com.client as win32
import win32con
import win32gui


def window_texts(hwnd: int) -> list[str]:
    """Return visible text from a top-level window and all child controls."""
    texts = [win32gui.GetWindowText(hwnd)]

    def collect(child_hwnd: int, _extra: object) -> None:
        """Collect text from one child control."""
        text = win32gui.GetWindowText(child_hwnd)
        if text:
            texts.append(text)

    win32gui.EnumChildWindows(hwnd, collect, None)
    return texts


def watch_for_compile_errors(
    stop_event: threading.Event,
    errors: list[str],
) -> None:
    """Close VBA compile dialogs and capture their diagnostic text."""
    pythoncom.CoInitialize()
    try:
        while not stop_event.is_set():
            windows: list[int] = []
            win32gui.EnumWindows(
                lambda hwnd, result: result.append(hwnd)
                if win32gui.IsWindowVisible(hwnd)
                else None,
                windows,
            )

            for hwnd in windows:
                texts = window_texts(hwnd)
                joined = " | ".join(texts)
                lowered = joined.lower()
                if "compile error:" in lowered or "syntax error" in lowered:
                    errors.append(joined)
                    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    stop_event.set()
                    return
            time.sleep(0.1)
    finally:
        pythoncom.CoUninitialize()


def main() -> int:
    """Compile the target workbook and return nonzero for any VBA error."""
    # Resolve the workbook path supplied by the diagnostic command.
    workbook_path = Path(sys.argv[1]).resolve()
    # Track compiler dialogs independently because Execute blocks on modal UI.
    stop_event = threading.Event()
    # Store any compiler error text found by the watcher.
    compile_errors: list[str] = []
    # Start an isolated Excel process for the compile probe.
    excel = win32.DispatchEx("Excel.Application")
    # Show Excel so VBA compile dialogs have a normal top-level owner.
    excel.Visible = True
    # Allow compile errors to display for diagnostic inspection.
    excel.DisplayAlerts = True
    # Disable workbook events so opening the file does not run unrelated code.
    excel.EnableEvents = False
    # Allow the embedded VBA project to be compiled.
    excel.AutomationSecurity = 1
    # Open the workbook without updating external links.
    workbook = excel.Workbooks.Open(
        str(workbook_path),
        UpdateLinks=0,
        ReadOnly=False,
        AddToMru=False,
    )
    # Show the VBA editor so the selected error line can be inspected.
    excel.VBE.MainWindow.Visible = True
    # Open the Debug menu from the VBA editor's menu bar.
    debug_menu = excel.VBE.CommandBars("Menu Bar").Controls(6)
    # Select the first Debug command, Compile VBAProject.
    compile_control = debug_menu.Controls(1)
    # Watch for a modal compile dialog before invoking the compiler.
    watcher = threading.Thread(
        target=watch_for_compile_errors,
        args=(stop_event, compile_errors),
        daemon=True,
    )
    watcher.start()

    try:
        # Execute only when the project still has pending code to compile.
        if compile_control.Enabled:
            compile_control.Execute()
        # Give the watcher a short final interval for delayed dialog creation.
        time.sleep(0.5)
    finally:
        stop_event.set()
        watcher.join(timeout=2)

    if compile_errors:
        print("VBA COMPILE FAILED", file=sys.stderr)
        print(compile_errors[0], file=sys.stderr)
        workbook.Close(SaveChanges=False)
        excel.Quit()
        return 1

    # Persist the successfully compiled VBA project in the staged workbook.
    workbook.Save()
    workbook.Close(SaveChanges=False)
    excel.Quit()
    print("VBA COMPILE PASSED")
    return 0


if __name__ == "__main__":
    # Run the diagnostic entry point.
    raise SystemExit(main())
