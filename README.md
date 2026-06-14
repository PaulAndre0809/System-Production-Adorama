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
| `Notes - Instructions` | Protected user and maintenance guide |
| `Dropdowns` | Hidden Campaign Type validation source |
| `Automation Log` | Hidden desktop automation and error log |

Monthly Calendar sheets and Last Week versus Current Week delivery comparisons
have been retired.

## Campaign Fields

Both campaign tables contain:

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

## Timed Link Labels

The `JIRA`, `ClickUp`, `Bluecore/Attentive`, and `Proof of Schedule` link
columns use native `HYPERLINK` formulas. They display the full URL until
exactly seven days after `Send Date` and numeric `Send Time`, then display the
clean platform name while preserving the same clickable URL.

For `STO`, `Local Timezone`, or a blank Send Time, the seven-day period starts
at midnight on the Send Date. Excel for the web updates the label when the
workbook recalculates; desktop Excel also schedules the next due refresh while
the workbook remains open.

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

The `Notes - Instructions` worksheet password is `adorama2024`. Protection is
an accidental-edit safeguard, not encryption.

Do not rename workbook tables or headers, expose or edit Dashboard helper
columns `AA:AL`, add blank lines after VBA continuation characters, or convert
the files to `.xlsx`.

See [AI Documentation Notes.md](AI%20Documentation%20Notes.md) for the detailed
technical architecture and QA process.
