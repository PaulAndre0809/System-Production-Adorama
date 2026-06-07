# Email & SMS Campaign Tracker

## Files

- `Email & SMS Campaign Tracker.xlsm`: working Excel application.
- `Email & SMS Campaign Tracker.bas`: canonical VBA source embedded in the XLSM.
- `Email & SMS Campaign Tracker.ps1`: transactional Windows deployment.
- `Email & SMS Campaign Tracker.py`: independent QA and maintenance utility.
- `Email & SMS Campaign Tracker.md`: operating and technical guide.

## Campaign sheets

`Email Campaigns` contains the existing email workflow. Its eight checkbox
columns remain in G:N:

1. Campaign Name and UTM Parameter (Source Code)
2. Creative Brief, SL & PH
3. SKUs
4. In-Design
5. Build, QA
6. Route
7. Approval
8. Segments

`SMS Campaigns` uses the same campaign metadata, links, delivery, and audit
fields, but has only four checkbox columns in G:J:

1. Send SMS Options
2. Send Test
3. Approval
4. Segments

Owner fields are plain text. Modern Microsoft 365 builds use native in-cell
checkboxes with TRUE/FALSE values. The VBA source uses late-bound checkbox
access so older desktop builds do not fail to compile; those builds receive a
visual Boolean fallback that can be toggled by double-clicking.

## Current Stage

Current Stage is a calculated table column on both campaign sheets. The
workbook is saved in automatic calculation mode, and desktop edit events also
calculate the changed row explicitly.

Email progresses through Source Code, Creative Brief, Waiting for SKUs, With
Design, Build / QA, Routing, Awaiting Approval, Segments, Links Pending, Ready
to Schedule, Scheduled, and Sent.

SMS progresses through SMS Options, Send Test, Awaiting Approval, Segments,
Links Pending, Ready to Schedule, Scheduled, and Sent.

Delivered values greater than zero set the stage to Sent.

## Dashboard and calendars

The Dashboard combines Email and SMS campaigns scheduled from Monday of the
current week through Sunday of the following week. A Channel column identifies
the source sheet.

`DeliveredComparisonTable` compares total delivered emails for:

- Last week, Monday through Sunday.
- Current week, Monday through Sunday.
- The current-week difference from last week.

All twelve calendar sheets combine both tables and prefix entries with
`Email |` or `SMS |`. The current-month tab is green and all other month tabs
use the default blue.

## SharePoint and Excel compatibility

The XLSM can be stored, opened, edited, and coauthored through SharePoint.
Microsoft documents native in-cell checkboxes for Microsoft 365, Mac, and Excel
for the web:

<https://support.microsoft.com/en-gb/office/using-check-boxes-in-excel-da85546d-c110-49b8-b633-9cebadcaf8d4>

Formula-driven stages, calendars, and delivered totals recalculate in supported
Excel for the web sessions. Microsoft does not allow VBA macros to run in Excel
for the web:

<https://support.microsoft.com/en-us/office/work-with-vba-macros-in-excel-for-the-web-98784ad0-898c-43aa-a1da-4f0fb5014343>

Consequently, desktop Excel is required for VBA-only conveniences: updater
name/timestamp stamping, the compact Dashboard detail snapshot, daily digest,
and automated deployment. Browser edits remain in the workbook and those
desktop-only views refresh the next time the file opens in desktop Excel.

Avoid simultaneous browser and desktop edits while running deployment.

## Deployment

Close all Excel windows, then run:

```powershell
powershell -ExecutionPolicy Bypass -File ".\Email & SMS Campaign Tracker.ps1"
```

The deployer:

1. Works on a temporary copy.
2. Keeps calculation manual during VBA replacement.
3. Compiles the VBA project.
4. Applies and validates both campaign models.
5. Calculates each worksheet once.
6. Saves the workbook in automatic calculation mode.
7. Replaces the XLSM only after QA succeeds.
8. Restores the previous XLSM if post-deployment QA fails.

## QA and maintenance

Run the disposable-copy QA suite:

```powershell
..\.venv\Scripts\python.exe ".\Email & SMS Campaign Tracker.py" --qa
```

Reapply the embedded configuration transactionally:

```powershell
..\.venv\Scripts\python.exe ".\Email & SMS Campaign Tracker.py" --apply
```

The suite validates both table schemas, checkbox behavior, dynamic Email and SMS
stage transitions, audit fields, calendar aggregation, Dashboard aggregation,
weekly delivered totals, automatic calculation persistence, save/reopen
persistence, embedded VBA, links, and broken references.

## Performance changes

- Removed persisted manual calculation, which caused stale Current Stage values.
- Removed broad `Application.CalculateFull` calls from calendar rebuilds.
- Uses targeted row and worksheet calculation.
- Keeps a compact 14-day Dashboard snapshot.
- Uses one transactional deployer instead of competing repair scripts.
- Applies native checkbox formatting by range instead of creating hundreds of
  shape controls.
