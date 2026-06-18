"""Bring the legacy backup's campaign tables in line with the active tracker:
add the missing "Scheduled" checkbox column and rebuild clean conditional formatting.

The backup's Email/SMS tables predate the "Scheduled" column, so they are one column
short and their schedule-gap highlighting points at the wrong column. The backup's
conditional formatting is also corrupted (stray `=#REF!` rules and a stranded
duplicate), which makes per-rule edits fail. For each table this tool:

  1. Inserts a "Scheduled" column immediately after "Segments" (matching the active
     layout), copying the sibling "Segments" column so it keeps the native checkbox
     style, then defaults every row to unticked (FALSE).
  2. Clears all conditional formatting on the sheet and rebuilds a clean set that
     matches the active tracker's intent (single-range, no duplicates, no #REF!):
       - schedule-gap highlight referencing the new Scheduled column,
       - five Current Stage rules, and
       - one Cancelled-row rule.

ValidateWorkbookConfiguration must return OK before saving. Idempotent.

Usage:
    python tools/fix_backup_scheduled_column.py <path> [<path> ...]
"""

import os
import shutil
import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com
import win32com.client as win32

XL_PASTE_ALL = -4104
ORANGE = 49407
YELLOW = 65535
# colours read from the active tracker's Current Stage / Cancelled rules
STAGE_RULES = [
    ('=NOT(ISERROR(SEARCH("Checked:",F2)))', 16247773),
    ('=F2="No checklist items checked"', 13431551),
    ('=F2="Sent"', 13561798),
    ('=F2="Scheduled"', 16247773),
    ('=F2="Ready to Schedule"', 10284031),
]
CANCELLED_COLOR = 10395391

WINDOW = ("IFERROR(INT($B2),0)>=TODAY()+1,"
          "IFERROR(INT($B2),0)<=TODAY()+IF(WEEKDAY(TODAY(),1)=6,3,1)")

TABLES = [
    dict(sheet="Email Campaigns", table="EmailCampaignsTable", lastcol="X",
         sched="$P", notes="$X", sg_color=ORANGE),
    dict(sheet="SMS Campaigns", table="SMSCampaignsTable", lastcol="S",
         sched="$L", notes="$S", sg_color=YELLOW),
]


def get_excel():
    try:
        return win32.gencache.EnsureDispatch("Excel.Application")
    except Exception:
        gp = getattr(win32com, "__gen_path__", None)
        if gp and os.path.isdir(gp):
            shutil.rmtree(gp, ignore_errors=True)
        return win32.gencache.EnsureDispatch("Excel.Application")


def retry_com(func, attempts=20, delay=0.5):
    last = None
    for i in range(attempts):
        try:
            pythoncom.PumpWaitingMessages()
            return func()
        except pywintypes.com_error as exc:
            last = exc
            if i == attempts - 1:
                raise
            time.sleep(delay)
    raise last


def col_names(lo):
    return [retry_com(lambda i=i: lo.ListColumns(i).Name)
            for i in range(1, retry_com(lambda: lo.ListColumns.Count) + 1)]


def add_expr_rule(ws, address, formula, color, stop, first=False):
    rng = retry_com(lambda: ws.Range(address))
    fc = retry_com(lambda: rng.FormatConditions.Add(Type=2, Formula1=formula))
    if first:
        retry_com(lambda: fc.SetFirstPriority())
    retry_com(lambda: setattr(fc, "StopIfTrue", stop))
    retry_com(lambda: setattr(fc.Interior, "Color", color))
    return fc


def fix_table(ws, cfg):
    lo = retry_com(lambda: ws.ListObjects(cfg["table"]))
    names = col_names(lo)
    if "Scheduled" not in names:
        seg_pos = names.index("Segments") + 1
        col = retry_com(lambda: lo.ListColumns.Add(seg_pos + 1))
        retry_com(lambda: setattr(col, "Name", "Scheduled"))
        seg = retry_com(lambda: lo.ListColumns("Segments").DataBodyRange)
        retry_com(lambda: seg.Copy())
        retry_com(lambda: col.DataBodyRange.PasteSpecial(Paste=XL_PASTE_ALL))
        retry_com(lambda: setattr(ws.Application, "CutCopyMode", False))
        retry_com(lambda: setattr(col.DataBodyRange, "Value", False))
        action = "inserted"
    else:
        action = "already present"

    db = retry_com(lambda: lo.DataBodyRange)
    last = retry_com(lambda: db.Row) + retry_com(lambda: db.Rows.Count) - 1
    data = f"A2:{cfg['lastcol']}{last}"
    stage = f"F2:F{last}"

    # nuke all (clears the #REF! corruption and stale duplicates) and rebuild cleanly
    retry_com(lambda: ws.Cells.FormatConditions.Delete())
    for formula, color in STAGE_RULES:
        add_expr_rule(ws, stage, formula, color, stop=True)
    add_expr_rule(
        ws, data,
        f'=OR(LOWER(TRIM({cfg["notes"]}2))="cancelled",LOWER(TRIM({cfg["notes"]}2))="canceled")',
        CANCELLED_COLOR, stop=False)
    sg = (f'=AND($D2<>"",{cfg["sched"]}2<>TRUE,' + WINDOW +
          f',NOT(OR(LOWER(TRIM({cfg["notes"]}2))="cancelled",LOWER(TRIM({cfg["notes"]}2))="canceled")))')
    fc = add_expr_rule(ws, data, sg, cfg["sg_color"], stop=False, first=True)
    retry_com(lambda: setattr(fc.Font, "Color", 0))
    print(f"    [{cfg['sheet']}] Scheduled {action}; CF rebuilt on {data} "
          f"(schedule-gap -> {cfg['sched']}, 5 stage, 1 cancelled)")


def process(excel, path: Path):
    print(f"  Opening {path.name}")
    wb = retry_com(lambda: excel.Workbooks.Open(str(path), UpdateLinks=0))
    try:
        name = wb.Name
        for cfg in TABLES:
            ws = retry_com(lambda c=cfg: wb.Worksheets(c["sheet"]))
            fix_table(ws, cfg)
        retry_com(lambda: excel.CalculateFull())
        validation = retry_com(lambda: excel.Run(f"'{name}'!ValidateWorkbookConfiguration"))
        print(f"    ValidateWorkbookConfiguration={validation!r}")
        if validation != "OK":
            raise RuntimeError(f"embedded validation returned {validation!r}")
        retry_com(wb.Save)
        print(f"  Saved {path.name}")
    finally:
        retry_com(lambda: wb.Close(False))


def main():
    if len(sys.argv) < 2:
        print("usage: fix_backup_scheduled_column.py <path> [<path> ...]")
        return 1
    paths = [Path(a).resolve() for a in sys.argv[1:]]
    for p in paths:
        if not p.exists():
            print(f"ERROR: missing {p}")
            return 1
    os.system("taskkill /F /IM EXCEL.EXE >nul 2>&1")
    time.sleep(1.0)
    excel = None
    failures = []
    try:
        excel = get_excel()
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1
        for p in paths:
            try:
                process(excel, p)
            except Exception as exc:  # noqa: BLE001
                failures.append((p.name, repr(exc)))
                print(f"  FAILED {p.name}: {exc!r}")
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
    if failures:
        print("\nCOMPLETED WITH ERRORS:")
        for n, e in failures:
            print(f" - {n}: {e}")
        return 1
    print("\nDONE")
    return 0


if __name__ == "__main__":
    sys.exit(main())
