# System Production Adorama

This repository contains the self-contained **Email & SMS Campaign Tracker**
Excel application.

## Workbooks

The `Production Tracker` directory contains:

- `Email & SMS Campaign Tracker.xlsm`: active tracker
- `Email & SMS Campaign Tracker Template.xlsm`: clean template
- `Email & SMS Campaign Tracker_backup.xlsm`: backup copy

All formulas, formatting, validation, and VBA required at runtime are embedded
in each `.xlsm` file. Scripts in `tools/` are development and QA utilities only.

## Worksheet Structure

| Sheet | Purpose |
| --- | --- |
| `Dashboard` | Current Sunday through next Saturday campaign view and summary KPIs |
| `Email Campaigns` | Email campaign source table and eight workflow checkboxes |
| `SMS Campaigns` | SMS campaign source table and four workflow checkboxes |
| `Notes - Instructions` | Plain-language, end-user guide to every feature (protected) |
| `June 2026 Calendar` | Monthly calendar mirrored from the SharePoint planning file |
| `May 2026 Calendar` | Hidden monthly calendar mirrored from SharePoint (active tracker only) |
| `Template for Duplicate` | Pre-formatted calendar used to create a new month |
| `Dropdowns` | Hidden Campaign Type validation source |
| `Automation Log` | Hidden desktop automation and error log |

The monthly Calendar sheets mirror the team's SharePoint Email/SMS planning
workbooks through external links. To add a month, copy `Template for Duplicate`,
rename it `<Month> 2026 Calendar`, and re-point its source via **Data > Edit
Links**. The Template and Backup copies ship with `June 2026 Calendar` +
`Template for Duplicate` as a guided example. The Last Week versus Current Week
delivery comparisons remain retired.

## Campaign Fields

Both campaign tables contain:

- `Week Number` (first column, left of `Send Date`; a `Week N` label worked out
  from the row's `Send Date` so campaigns can be grouped and filtered by week)
- `Send Date`
- `Send Time`
- `Campaign Name`
- `Campaign Type`
- `Current Stage`
- `Owner`
- channel-specific workflow checkboxes
- links
- `Est. Audience`
- `Delivered`
- `Last Updated`
- `Last Updated By`
- `Notes`

Email workflow fields are:

1. `Campaign Name and UTM Parameter (Source Code)`
2. `Creative Brief, SL & PH`
3. `SKUs`
4. `In-Design`
5. `Build, QA`
6. `Route`
7. `Approval`
8. `Segments`

SMS workflow fields are:

1. `Send SMS Options`
2. `Send Test`
3. `Approval`
4. `Segments`

`Current Stage` lists all checked workflow fields rather than selecting only one
stage.

## Date And Time Input

- `Send Date` displays as `dddd, mmmm d, yyyy`, for example
  `Wednesday, June 10, 2026`.
- Numeric `Send Time` values display in 12-hour format.
- `Send Time` also accepts text such as `STO` and `Local Timezone`.

## Dashboard

The Dashboard combines Email and SMS campaigns scheduled from the current
Sunday through the following Saturday.

Approval displays `Done` or `Not Yet`. Segments displays `Provided` or
`Pending`. A campaign is excluded from both the Dashboard feed and summary
KPIs when either `Current Stage` or `Notes` is exactly `Cancelled` or
`Canceled`, ignoring capitalization and surrounding spaces.

Summary KPIs use expanding structured table references:

- Active Work
- Sending Today
- Email Active
- SMS Active
- Approval Pending
- Sent

A `Week Number` tile beside the KPIs shows the current two-week window as a span
(for example `25-26`), matching the Sunday-through-next-Saturday feed. Beneath it,
a per-campaign `Week Number` column (next to `Bluecore/Attentive`) shows a
`Week N` label for each row, derived from the leading `MMDDYY` code in the
campaign string (e.g. `061726-STO-...` → `Week 25`) and falling back to the Send
Date when there is no code.

## Timed Link Labels

The `JIRA`, `ClickUp`, `Bluecore/Attentive`, and `Proof of Schedule` link
columns use native `HYPERLINK` formulas. They display the full URL until
exactly seven days after `Send Date` and numeric `Send Time`, then display the
clean platform name while preserving the same clickable URL.

For `STO`, `Local Timezone`, or a blank Send Time, the seven-day period starts
at midnight on the Send Date. Excel for the web updates the label when the
workbook recalculates; desktop Excel also schedules the next due refresh while
the workbook remains open.

## Schedule-Gap Highlighting

On the `Email Campaigns` and `SMS Campaigns` sheets, a row is highlighted when a
deployment is due in the next working window but its `Scheduled` checkbox is still
unchecked. Email rows fill orange and SMS rows fill yellow. The window is the next
day, and on Fridays it extends through Saturday, Sunday, and Monday so weekend and
Monday sends are flagged before everyone leaves. Cancelled rows are never
highlighted, and checking the `Scheduled` box clears the highlight. The rule is
native conditional formatting, so it updates with the date and works in desktop
Excel and Excel for the web.

The same highlighting also appears on the **Dashboard** feed. Because the
Dashboard has no `Scheduled` column, a feed row is treated as unscheduled when
its `Stage` is not yet `Scheduled` or `Sent`; Email rows fill orange and SMS rows
fill yellow over the same next-day/Friday-weekend window.

## Excel And SharePoint Compatibility

Native formulas, tables, filters, formats, and saved checkbox values work in
desktop Excel and Excel for the web.

VBA does not execute in Excel for the web. Open the workbook in desktop Excel
with macros enabled for:

- automatic audit timestamps and editor names
- event-driven formatting repair
- checkbox double-click fallback
- Dashboard refresh commands
- automation logging

Use SharePoint version history as the authoritative editor record for web
changes.

## Maintenance

The `Notes - Instructions` sheet inside each workbook is the **end-user guide**:
plain, non-technical wording (a *Feature / What you do / How it works / Please do /
Please avoid / Where it works* table) covering only features still in use. This
README and `AI Documentation Notes.md` are the technical/maintainer references.

The `Notes - Instructions` worksheet password is `Adorama@042026_` for the active
tracker and template; the legacy backup copy still opens with `adorama2024`.
Protection is an accidental-edit safeguard, not encryption.

Do not rename workbook tables or headers, expose or edit Dashboard helper
columns `AA:AL`, add blank lines after VBA continuation characters, or convert
the files to `.xlsx`.

See [AI Documentation Notes.md](AI%20Documentation%20Notes.md) for the detailed
technical architecture and QA process.
