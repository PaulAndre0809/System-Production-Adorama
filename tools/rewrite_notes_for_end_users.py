"""Rewrite the Notes - Instructions sheet in plain, end-user-friendly language.

Same five-column layout and the same active features, but every cell is reworded
for a non-technical reader (no VBA/structured-reference/helper-column jargon) while
staying precise. References only features that are still in use. The title in A1
("Detailed Notes and Instructions") and the wording that QA checks for
("Wednesday, June 10, 2026", "STO or Local Timezone", "Monthly Calendars",
"Timed Link Labels", "Cancelled Campaigns") are preserved.

Usage:
    python tools/rewrite_notes_for_end_users.py [<path> ...]   # default: all three
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

NOTES_PASSWORDS = ["Adorama@042026_", "adorama2024"]

HEADER = ("Feature", "What you do", "How it works (in plain words)",
          "Please do / Please avoid", "Where it works")

# rows 4..23 : (Feature, What you do, How it works, Do/Avoid, Where it works)
ROWS = [
    ("Email Campaigns", "Add or change an email campaign",
     "This is the list of every email campaign. Put one campaign on each row. The most "
     "important box is the Send Date - that is the day the email goes out, and the Dashboard "
     "summary page uses it to show what is coming up.",
     "Please fill in at least the Campaign Name and the Send Date. Do not rename the column "
     "titles along the top, and do not delete the columns that fill in by themselves.",
     "Works in the Excel app on your computer and in a web browser."),

    ("Email Campaigns", "Tick off each step as you finish it",
     "Every campaign has a row of small tick boxes for the work that needs doing. Click a box "
     "to tick it. The Current Stage box then lists everything you have ticked, and says "
     "Completed once all the needed steps are done.",
     "Please use the tick boxes that are already there. Do not type words in them or swap them "
     "for drop-down lists.",
     "On your computer you click to tick. In a web browser, type TRUE to tick or FALSE to "
     "untick."),

    ("SMS Campaigns", "Add or change a text-message campaign",
     "This list works exactly like the email list, but for text messages. Its steps are Send "
     "SMS Options, Send Test, Approval, and Segments.",
     "Please fill in the Campaign Name and the Send Date. Do not rename this sheet or its "
     "column titles.",
     "The Dashboard blends the email and text campaigns together for you automatically."),

    ("Campaign Type", "Pick the kind of campaign",
     "Click the little arrow in the Campaign Type box and choose from the list: Promo, "
     "Services, Loyalty, PLCC, Newsletters, Events, NPA, Others, or leave it blank. Loyalty "
     "and PLCC are now two separate choices. If you pick Others, a small box pops up so you "
     "can type your own wording for just that one row.",
     "The list of choices is kept on a hidden helper sheet called Dropdowns. Please do not "
     "delete that sheet.",
     "Anything you type into the Others pop-up is used on that one row only - it is not added "
     "to the list for next time."),

    ("Filters", "Temporarily hide rows you don't need",
     "Click the little arrow at the top of a column (for example Send Date or Delivered) to "
     "hide some rows for now, so you can focus on, say, just this month's work.",
     "If a campaign looks like it has gone missing, clear the filters first - it is almost "
     "always just hidden.",
     "Works in the Excel app on your computer and in a web browser."),

    ("Dashboard", "See what is coming up",
     "The Dashboard is the summary page. Its main table gathers the email and text campaigns "
     "going out from this Sunday through the following Saturday, all in one place.",
     "Please make your changes on the Email or SMS sheets, not on the Dashboard - the "
     "Dashboard fills itself in. (Approval shows Done or Not Yet; Segments shows Provided or "
     "Pending.)",
     "You can read it in a web browser; the full automatic updating happens in the Excel app "
     "on your computer."),

    ("Dashboard", "Update the Dashboard",
     "The Last Refresh time updates on its own. If you have just changed the campaign lists "
     "and want the Dashboard to catch up straight away, use the Refresh Dashboard command "
     "(in the Excel app on your computer).",
     "Off to the right are some light-gray helper columns the file uses to do its sums. "
     "Please leave them hidden and do not type in them.",
     "A web browser cannot run the little helper programs, but the saved numbers and tables "
     "still show correctly."),

    ("Monthly Calendars", "Look at the planning calendars",
     "The June 2026 Calendar (and a hidden May 2026 Calendar) show the team's monthly plan. "
     "They fill in automatically from the team's files kept on SharePoint, so you do not type "
     "anything here. Template for Duplicate is a blank, ready-made month you can copy to make "
     "a new one.",
     "To add a new month: (1) right-click the Template for Duplicate tab and choose Move or "
     "Copy to make a copy; (2) rename the copy, for example July 2026 Calendar; (3) go to "
     "Data > Edit Links and point it to that month's SharePoint files; (4) click Update. "
     "Please keep the calendar's look exactly the same.",
     "The calendars refresh in the Excel app on your computer when you can reach the "
     "SharePoint files - if asked, click Update Values."),

    ("Audit Fields", "See who changed a row, and when",
     "The Last Updated and Last Updated By boxes quietly record the date and time, and the "
     "name of the person who last changed a campaign row. On your computer this fills in by "
     "itself.",
     "Please do not type over these boxes. (In a web browser you may fill in Last Updated By "
     "yourself if you wish.)",
     "A web browser cannot record the name automatically - to see who changed something "
     "there, use SharePoint's version history."),

    ("SharePoint", "Share the file safely",
     "Keep this file in SharePoint and open it from there, so everyone is working on the same "
     "copy.",
     "Please do not save it as a plain .xlsx file - that throws away the built-in helpers. "
     "Try not to have two people rearranging the tables at the very same moment.",
     "The automatic helpers run in the Excel app on your computer; the numbers, tables, and "
     "filters can still be viewed in a web browser."),

    ("Maintenance", "Keep everything tidy",
     "If the Dashboard ever looks out of date, use the Refresh Dashboard command. It simply "
     "re-does the sums; it does not rebuild or change your campaign lists.",
     "Before making big changes to the tables or the column titles, save a backup copy (or a "
     "SharePoint version) first, just in case.",
     "A refresh repairs the Dashboard's sums and look without touching your source lists or "
     "the calendars."),

    ("Send Date", "Type the day a campaign goes out",
     "Type a normal date, for example 6/10/2026. The file then shows it nicely as "
     "Wednesday, June 10, 2026, but underneath it stays a real date the Dashboard can use.",
     "Please type a real date. Do not type the weekday and month out as a plain sentence.",
     "The nice Wednesday, June 10, 2026 style is saved in the file and looks the same on your "
     "computer and in a web browser."),

    ("Send Time", "Type a time, or a short label",
     "Type a normal time and it shows as 10:00 AM or 10:00 PM. When there is no exact time "
     "yet, you may instead type a short label such as STO or Local Timezone.",
     "Please type either a real time or the agreed label. Do not add a rule that blocks "
     "typing in this box.",
     "Words stay as words; real times automatically get the 10:00 AM style."),

    ("Cancelled Campaigns", "Take a campaign off the Dashboard",
     "To remove a campaign from the Dashboard, type just the single word Cancelled (or "
     "Canceled) in its Notes box - nothing else.",
     "Type only that one word. Extra words such as 'Cancelled by team' will NOT count "
     "(capital letters and spaces do not matter, but extra words do).",
     "The Dashboard list and the summary count boxes both follow this same rule."),

    ("Timed Link Labels", "Links tidy themselves up after a week",
     "For the JIRA, ClickUp, Bluecore/Attentive, and Proof of Schedule links, the box shows "
     "the full web address for the first seven days after the send. After that it neatly shows "
     "just the name (such as JIRA), and the link still works when clicked.",
     "For exact timing, give the row a real Send Date and a real time. If the time is STO, "
     "Local Timezone, or left blank, the seven days are counted from midnight on the Send "
     "Date.",
     "Works in the Excel app on your computer and in a web browser. Your computer also "
     "refreshes it at the next due time while the file is open."),

    ("Troubleshooting", "If something does not look right",
     "Most of the automatic features only run in the Excel app on your computer, with macros "
     "(the built-in helper programs) allowed. If something seems stuck, close the file and "
     "open it again on your computer.",
     "When in doubt, open the file in the Excel app on your computer and click Enable if it "
     "asks about macros.",
     "A web browser shows the saved numbers and formatting but does not run the helpers."),

    ("Performance", "The file stays quick on its own",
     "The file is set up to stay fast. It updates the Dashboard once after a batch of edits, "
     "rather than after every single change.",
     "Please do not rename the tables or column titles, and do not paste over the hidden "
     "helper area - that is what keeps it fast and correct.",
     "The Refresh Dashboard command only re-does the Dashboard's sums, so it stays speedy."),

    ("Sheet Protection", "This page is locked on purpose",
     "This Notes page is locked so it cannot be changed by accident.",
     "If the person who looks after the file needs to edit it, the unlock password is "
     "Adorama@042026_. Please lock the page again afterwards.",
     "The lock only prevents accidental edits - it is not security and does not hide the "
     "file's contents."),

    ("Schedule-Gap Highlighting", "A colour warning for sends that are not booked yet",
     "On the Email and SMS sheets, a row turns orange (email) or yellow (text) when it is due "
     "to go out the next day but its Scheduled box is not ticked yet. On Fridays the warning "
     "also covers Saturday, Sunday, and Monday. The Dashboard shows the same warning colours.",
     "Once the campaign is booked, tick the Scheduled box and the colour clears by itself. "
     "Please do not remove this colouring or rename the Scheduled, Send Date, Campaign Name, "
     "or Notes columns.",
     "The colouring updates with today's date and works on your computer and in a web "
     "browser."),

    ("Dashboard Week Number", "See which week of the year it is",
     "Beside the summary boxes, a small Week Number tile shows the current two weeks of the "
     "year as a range, such as 25-26. The campaign list also has a Week Number column that "
     "shows the week for each row (it reads the date code at the start of the campaign name, "
     "or uses the Send Date when there is no code).",
     "This is for reading only - please do not type over it.",
     "It updates by itself with today's date, on your computer and in a web browser."),
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


def rewrite_notes(wb):
    ws = retry_com(lambda: wb.Worksheets("Notes - Instructions"))
    used_pw = None
    for pw in NOTES_PASSWORDS:
        try:
            retry_com(lambda pw=pw: ws.Unprotect(Password=pw))
            used_pw = pw
            break
        except pywintypes.com_error:
            continue
    if used_pw is None:
        raise RuntimeError("could not unprotect Notes - Instructions")
    try:
        for c, val in enumerate(HEADER, start=1):
            retry_com(lambda c=c, val=val: setattr(ws.Cells(3, c), "Value", val))
        for i, row in enumerate(ROWS):
            r = 4 + i
            for c, val in enumerate(row, start=1):
                retry_com(lambda r=r, c=c, val=val: setattr(ws.Cells(r, c), "Value", val))
        retry_com(lambda: ws.Range(f"A3:E{3 + len(ROWS)}").EntireRow.AutoFit())
        print(f"    Rewrote header + {len(ROWS)} rows in plain language")
    finally:
        retry_com(lambda: ws.Protect(Password=used_pw, DrawingObjects=True, Contents=True, Scenarios=True))


def process(excel, path: Path):
    print(f"  Opening {path.name}")
    wb = retry_com(lambda: excel.Workbooks.Open(str(path), UpdateLinks=0))
    try:
        name = wb.Name
        rewrite_notes(wb)
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
    repo = Path(__file__).resolve().parents[1]
    folder = repo / "Production Tracker"
    paths = ([Path(a).resolve() for a in sys.argv[1:]] if len(sys.argv) > 1
             else [folder / "Email & SMS Campaign Tracker.xlsm",
                   folder / "Email & SMS Campaign Tracker Template.xlsm",
                   folder / "Email & SMS Campaign Tracker_backup.xlsm"])
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
