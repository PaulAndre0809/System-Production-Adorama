# AI Documentation Notes

## Current System

The three `Email & SMS Campaign Tracker` workbooks are self-contained `.xlsm`
files. Runtime behavior does not depend on Python, PowerShell, BAS, or other
external files. Development scripts under `tools/` are used only to build and
test releases.

The current workbook architecture uses two source tables:

- `EmailCampaignsTable` on `Email Campaigns`
- `SMSCampaignsTable` on `SMS Campaigns`

The remaining user-facing sheets are `Dashboard` and
`Notes - Instructions`. `Dropdowns` and `Automation Log` are hidden support
sheets.

Monthly Calendar sheets and the former Last Week versus Current Week delivery
comparisons were intentionally retired. Compatibility macros with old calendar
names are safe no-ops and must not recreate those features.

## Active VBA Entry Points

### `Workbook_Open`

- Removes frozen or split views.
- Restores campaign date and time formats.
- Repairs and recalculates Dashboard formulas.
- Keeps startup work narrowly scoped for performance.

### `HandleCampaignChange`

- Runs from the Email and SMS `Worksheet_Change` events in desktop Excel.
- Prompts (via `PromptCustomCampaignType`/`InputBox`) for a custom value when a
  `Campaign Type` cell is set to `Others`, writing the entry to that row only.
- Restores `Send Date` and `Send Time` formatting after typing or pasting.
- Converts supported pasted URLs to native timed hyperlinks.
- Updates `Last Updated` and `Last Updated By`.
- Recalculates affected `Current Stage` cells.
- Auto-fits changed rows as one edit batch.
- Refreshes Dashboard outputs once per edit batch.
- Logs unexpected failures to `Automation Log`.

Excel for the web does not execute VBA events. SharePoint version history is
the authoritative editor record for web edits.

### `ToggleInventoryChecklist`

- Toggles supported workflow cells between Boolean `TRUE` and `FALSE`.
- Modern Excel uses native in-cell checkbox controls.
- Older desktop versions use the workbook's visual Boolean fallback.

### `ApplyCampaignEntryFormats`

- Applies `dddd, mmmm d, yyyy` to `Send Date`.
- Applies `h:mm AM/PM` to numeric `Send Time` values.
- Removes blocking validation from `Send Time`, allowing labels such as `STO`
  and `Local Timezone`.
- Applies matching Dashboard date and time presentation.

### `ApplyDashboardKpiFormulas`

Repairs the six KPI cells with expanding table-column formulas:

| KPI | Definition |
| --- | --- |
| Active Work | Email Active plus SMS Active |
| Sending Today | Non-cancelled campaigns scheduled today |
| Email Active | Email campaigns with no delivered total, excluding cancelled rows |
| SMS Active | SMS campaigns with no delivered total, excluding cancelled rows |
| Approval Pending | Non-cancelled campaigns where Approval is `FALSE` |
| Sent | Campaigns where Delivered is greater than zero |

These formulas use full structured references. Row-scoped references such as
`[@Delivered]` are invalid for Dashboard KPI cells and must not be reintroduced.

### `RefreshDashboard`

- Repairs KPI formulas.
- Recalculates timed source-link labels.
- Calculates the hidden native spill formula at `Dashboard!AA11`.
- Recalculates `DashboardWorkTable`, KPI cells, and audit display cells.
- Does not rebuild source tables or removed Calendar sheets.

### `ValidateWorkbookConfiguration`

Performs lightweight embedded validation of core sheets, table structure,
retired-feature removal, and instruction-sheet availability. External QA adds
VBA compilation, seeded editing scenarios, formula verification, and package
integrity checks.

## Formula Architecture

### Current Stage

`Current Stage` is formula-driven on both source sheets. It displays one status
containing every checked workflow field:

- `No checklist items checked`
- `Checked: <all selected workflow columns>`

Email checks eight workflow fields. SMS checks `Send SMS Options`, `Send Test`,
`Approval`, and `Segments`.

### Dashboard Campaign Window

The hidden `Dashboard!AA11` formula uses `LET`, `FILTER`, `HSTACK`, `VSTACK`,
and `SORTBY` to combine Email and SMS campaigns scheduled from the current
Sunday through the following Saturday.

Rows are excluded when either `Current Stage` or `Notes` is exactly
`Cancelled` or `Canceled`, ignoring capitalization and surrounding spaces.
Partial phrases are not treated as the cancellation status. Dashboard Approval
displays `Done` or `Not Yet`; Segments displays `Provided` or `Pending`.

### Timed Link Labels

`Jira Link`, `ClickUp Link`, `Bluecore/Attentive Link`, and
`Proof of Schedule` store their URLs inside native `HYPERLINK` formulas. The
formula displays the full URL before the maturity timestamp and `JIRA`,
`ClickUp`, `Bluecore/Attentive`, or `Proof of Schedule` after it.

The maturity timestamp is `Send Date + numeric Send Time + 7 days`. Text or
blank Send Time values use midnight on Send Date. Desktop Excel schedules an
`Application.OnTime` recalculation for the next maturity while the workbook is
open. Excel for the web cannot run VBA, but the native `NOW()` formula updates
when the workbook recalculates.

### Audit Display

- `Last Refresh` is formula-driven and uses 12-hour time.
- `Last Edited By` reads the user associated with the newest desktop audit
  timestamp.
- Web users must rely on SharePoint version history because workbook formulas
  cannot reliably retrieve the signed-in web editor.

### Schedule-Gap Highlighting

`Email Campaigns` and `SMS Campaigns` carry a native conditional-formatting rule
that flags deployments due soon but not yet scheduled. An Email row fills orange
(`#FFC000`) and an SMS row fills yellow (`#FFFF00`) when every condition holds:

- `Campaign Name` is not blank.
- `Scheduled` is not `TRUE` (Email column `O`, SMS column `K`).
- `Send Date` falls inside the next working window.
- `Notes` is not exactly `Cancelled` or `Canceled`.

The window is `TODAY()+1` through `TODAY()+IF(WEEKDAY(TODAY(),1)=6,3,1)`, so most
days highlight only the next day, while Fridays extend through Saturday, Sunday,
and Monday. The whole-row rule is set to first priority and applies to
`Email Campaigns!A2:W207` and `SMS Campaigns!A2:R201`. The Email formula is:

```
=AND($C2<>"",$O2<>TRUE,IFERROR(INT($A2),0)>=TODAY()+1,
  IFERROR(INT($A2),0)<=TODAY()+IF(WEEKDAY(TODAY(),1)=6,3,1),
  NOT(OR(LOWER(TRIM($W2))="cancelled",LOWER(TRIM($W2))="canceled")))
```

The SMS formula is identical except `$K2` replaces `$O2` and `$R2` replaces
`$W2`. Because the rule is native conditional formatting driven by `TODAY()`, it
recalculates daily and works in Excel for the web. It coexists with the existing
Cancelled and Current Stage rules and survives `RefreshDashboard` and
`RefreshNativeOutputs`. The `IFERROR(INT(...),0)` guard ignores text or blank
`Send Time`/`Send Date` values, mirroring the Dashboard campaign-window formula.

### Dashboard Week Number

`Dashboard!M4:N6` is a KPI-style tile labeled `Week Number`, placed to the right
of the `Sent` tile and styled to match it. The value cell `M5` holds:

```
=WEEKNUM(TODAY(),1)&"-"&WEEKNUM(TODAY()+7,1)
```

It shows the current two-week window as a span, for example `25-26` for the week
of June 14, 2026. `WEEKNUM(..,1)` uses a Sunday-start week so the number matches
the Sunday-through-next-Saturday Dashboard feed, and the second term reports the
following week. The tile is display-only and persists across Dashboard refreshes.

A per-campaign `Week Number` column sits beneath that tile at `Dashboard!M`,
immediately right of the `DashboardWorkTable` `Bluecore/Attentive` column. It is a
standalone sheet column (header `M10`, formula `M11:M160`), deliberately **not** a
13th `DashboardWorkTable` column, because `ValidateWorkbookConfiguration` requires
that table to keep exactly 12 columns. Each row shows a `Week N` label derived
from the leading `MMDDYY` code in the `Campaign` value, e.g.
`061726-STO-Services-Trade-P-B-NA-GLP` resolves to `Week 25`. When the campaign
has no parseable code (such as `TBD`), it falls back to the row's `Send Date`;
blank feed rows stay blank. The first-data-row formula is:

```
=LET(sendDate,$A11,camp,$D11&"",code,LEFT(camp,6),
  parsed,IFERROR(DATE(2000+VALUE(MID(code,5,2)),VALUE(LEFT(code,2)),VALUE(MID(code,3,2))),""),
  useDate,IF(parsed="",IF(ISNUMBER(sendDate),INT(sendDate),""),parsed),
  IF(useDate="","","Week "&WEEKNUM(useDate,1)))
```

The column uses `WEEKNUM(..,1)` for the same Sunday-start week as the tile and
survives `RefreshDashboard`/`RefreshNativeOutputs` (which leave columns outside the
table untouched). Workbook auto-expand is left disabled when it is written so the
table does not absorb it.

## Data Entry Rules

- `Send Date` must be a real Excel date and displays like
  `Wednesday, June 10, 2026`.
- `Send Time` accepts real Excel times and text such as `STO` or
  `Local Timezone`.
- `Campaign Type` offers Promo, Services, Loyalty, PLCC, Newsletters, Events,
  NPA, Others, and blank. `Loyalty` and `PLCC` are independent options (the former
  combined `Loyalty & PLCC` option was split). In desktop Excel, selecting `Others`
  opens a pop-up (`InputBox`) to type a custom campaign type that fills that row
  only. Custom text is allowed but is not added to the dropdown list. The list
  source is `CampaignTypeOptions()` in VBA, mirrored to `Dropdowns!A2:A10`.
- `Owner` and `Notes` are plain text.
- Workflow columns contain only Boolean checkbox values.
- Do not rename tables, calculated columns, or headers.

## Protection And Compatibility

`Notes - Instructions` is protected against accidental edits. The maintenance
password is `Adorama@042026_` for the active tracker and template (verified
against the sheet's stored SHA-512 protection hash). The older backup copy still
opens with the legacy `adorama2024`. Worksheet protection is not encryption.

Native formulas, tables, filters, saved formatting, and checkbox values work in
Excel for the web. VBA compilation, edit events, automatic audit stamping, and
macro commands require desktop Excel with macros enabled.

## VBA Compile Integrity

VBA line continuation characters (`_`) must be followed immediately by the
next physical code line. A blank line after `_` causes a compile-time syntax
error. The release builder removes these invalid gaps from every VBA component,
and QA rejects any recurrence.

## Performance Guidance

- Use structured table references instead of whole worksheet columns.
- Recalculate Dashboard ranges once per edit batch.
- Keep helper columns `AA:AL` hidden and intact.
- Avoid volatile formulas except the intentional `TODAY()` and `NOW()` audit
  and scheduling calculations.
- Do not run retired migration or calendar rebuild routines.

## QA Release Process

Each release is checked for:

1. XLSM package and embedded VBA integrity.
2. VBA continuation syntax and full VBA project compilation.
3. Preservation of existing campaign data and formulas.
4. Ten temporary Email and ten temporary SMS campaign scenarios.
5. Checkbox, Current Stage, validation, filtering, date, time, and audit logic.
6. Dashboard date window, friendly statuses, cancellation exclusion, and KPIs.
7. Exact seven-day hyperlink behavior before and after the maturity timestamp.
8. Formula errors, broken names, freeze panes, and external workbook links.
9. Instruction-sheet protection and password verification.
10. Schedule-gap highlighting rules and the Dashboard Week Number tile survive a
    Dashboard refresh, and the highlight rules add no formula errors.

Temporary QA records are created only in disposable copies.
