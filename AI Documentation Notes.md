# AI Documentation Notes

## Current System

The three `Email & SMS Campaign Tracker` workbooks are `.xlsm` files. Their
campaign tracking, Dashboard, and VBA do not depend on Python, PowerShell, BAS,
or other external files; development scripts under `tools/` are used only to
build and test releases. The monthly Calendar sheets are the one exception: they
mirror the team's SharePoint planning files through external workbook links (see
*Monthly Calendars* below).

The current workbook architecture uses two source tables:

- `EmailCampaignsTable` on `Email Campaigns`
- `SMSCampaignsTable` on `SMS Campaigns`

The remaining user-facing sheets are `Dashboard`, `Notes - Instructions`, the
SharePoint-linked monthly calendars (`June 2026 Calendar`, hidden
`May 2026 Calendar`), and `Template for Duplicate`. `Dropdowns` and
`Automation Log` are hidden support sheets.

The former Last Week versus Current Week delivery comparisons remain retired (do
not recreate `DeliveredComparison` objects). The monthly Calendar sheets are no
longer retired — they are an active feature that mirrors SharePoint planning
files. The legacy VBA calendar routines (`RebuildMonthlyCalendars`,
`RemoveLegacyCalendarAndComparisonArtifacts`) are unused and must not be run:
they build or delete differently-named sheets and would not respect the
SharePoint mirror.

## Repository Layout and Source of Truth

**Purpose:** Define which files are authoritative and which are disposable
artifacts, so future agents do not edit or trust the wrong copy.

| Path | Role | Authoritative? |
| --- | --- | --- |
| `Production Tracker/*.xlsm` | Active tracker, Template, and legacy Backup workbooks. Embed all runtime formulas, formatting, validation, checkbox state, and the live VBA project. | **Yes** — the workbooks are the product. |
| `Project Tracker/Loyalty and PLCC Email Plan Template.xlsx` | Standalone planning template (no macros). | Yes (self-contained). |
| `tools/*.py` | Development, build, patch, and QA utilities. Not loaded at runtime. | Yes, as the build pipeline. |
| `vba_dump.txt`, `clean_temp*.vba`, `clean_mod.vba`, `clean.vba`, `notes_dump.txt`, `tools/vba_dump.txt` | Point-in-time VBA/text **exports** captured during development. | **No** — historical snapshots only. |
| `Reporting Analysis/` | Reserved/empty working area. | n/a |

**Critical insight (verified by static analysis):** the loose VBA dumps are
**stale**. The root `vba_dump.txt` is the standard module only and `tools/vba_dump.txt`
adds the sheet/workbook event modules, but **neither** contains the
`ApplyCampaignEntryFormats`, `ApplyDashboardKpiFormulas`, or
`PromptCustomCampaignType` procedures. Those procedures exist in the **live
workbook**, injected by `tools/fix_campaign_entry_formats.py` and
`tools/split_campaign_type_and_others_prompt.py` after the dumps were taken. To
read the current VBA, export it from the `.xlsm` (`tools/export_vba.py`) rather
than reading the loose dumps.

## VBA Module Map

The live VBA project is organized as event modules that delegate to a single
standard module (`modEmailProductionTracker`).

### Event modules (thin delegators)

| Module | Handler | Delegates to |
| --- | --- | --- |
| `ThisWorkbook` | `Workbook_Open` | Startup repair (unfreeze, restore formats, repair/recalc Dashboard) |
| `Email Campaigns` sheet | `Worksheet_Change` | `HandleInventoryChange` → `HandleCampaignChange` |
| `Email Campaigns` sheet | `Worksheet_BeforeDoubleClick` | `ToggleInventoryChecklist` (legacy checkbox fallback) |
| `SMS Campaigns` sheet | `Worksheet_Change` | `HandleCampaignChange` |
| `SMS Campaigns` sheet | `Worksheet_BeforeDoubleClick` | `ToggleInventoryChecklist` |

Excel for the web never fires these events; SharePoint version history is the
authoritative web editor record.

### `modEmailProductionTracker` — core procedure catalog

Signatures use the live VBA. Procedures marked *(injected)* are added to the
workbook by `tools/` scripts and are absent from the loose dumps.

| Procedure | Signature → Return | Purpose |
| --- | --- | --- |
| `GenerateCampaignID` | `(sendDate As Variant, campaignName As String) As String` | Build a stable campaign source code from date + cleaned name. |
| `CalculateCurrentStage` | `(ws As Worksheet, r As Long) As String` | Compose the `Current Stage` text listing every checked workflow field. |
| `CalculateRiskLevel` | `(ws As Worksheet, r As Long) As String` | Derive a risk label from send date proximity and checklist completion. |
| `ToggleInventoryChecklist` | `(Target As Range) As Boolean` | Flip a workflow cell TRUE/FALSE; native checkbox first, legacy fallback otherwise. |
| `HandleCampaignChange` | `(ws As Worksheet, Target As Range)` | Master `Worksheet_Change` handler: Others prompt, format repair, URL→hyperlink, audit stamp, stage recalc, batched Dashboard refresh, error logging. |
| `PromptCustomCampaignType` *(injected)* | `(lo As ListObject, changedCells As Range)` | `InputBox` prompt when `Campaign Type` = `Others`; writes the custom value to that row only. |
| `UpdateRowTimestampAndUser` | `(rowNumber As Long, lo As ListObject)` | Stamp `Last Updated` / `Last Updated By` for an edited row. |
| `ApplyCampaignEntryFormats` *(injected)* | `()` | Apply long-date / 12-hour formats and remove blocking `Send Time` validation. |
| `ApplyDashboardKpiFormulas` *(injected)* | `()` | Repair the six KPI cells with expanding structured-reference formulas. |
| `RefreshDashboard` | `()` | Repair KPI formulas, recalc timed link labels, the `AA11` spill, the work table, and audit cells. |
| `RefreshNativeOutputs` | `()` | Recalculate native spill/feed outputs without rebuilding source tables. |
| `ValidateWorkbookConfiguration` | `() As String` | Embedded self-check; returns `"OK"` or a failure description (the QA gate's first assertion). |
| `InventoryColumnNumber` | `(headerName As String) As Long` | Resolve a campaign column index by header name (VBA never hard-codes column letters). |
| `LogAction` | `(actionName As String, details As String)` | Append a row to the hidden `Automation Log`. |
| `CreateDailyDigest` | `()` | Build the daily digest summary. |
| `UnfreezeWorkbookViews` | `(wb As Workbook)` | Remove freeze panes / split views (called on open). |

**Build/config procedures (run only during release assembly, never at runtime):**
`MigrateProductionInventoryStructure`, `ApplyAllConfigurations`,
`RefreshProductionStatus`, `UpdateCalendarTabs`, `CreateBackupCopy`,
`BuildNotesInstructionSheet` (+ `AddInstructionRow`), and the styling helpers.

**Do-not-run legacy procedures:** `RebuildMonthlyCalendars`,
`RemoveLegacyCalendarAndComparisonArtifacts`, `CreateDeliveredComparison`,
`CreateDeliveredComparisonChart` — they build/delete differently-named sheets and
ignore the SharePoint calendar mirror.

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
- Does not rebuild source tables or the SharePoint-linked Calendar sheets.

### `ValidateWorkbookConfiguration`

Performs lightweight embedded validation of core sheets, table structure,
retired-comparison removal, and instruction-sheet availability. It no longer
flags monthly Calendar sheets. External QA adds
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

**Long URLs:** a string literal longer than 255 characters cannot live inside a
formula — Excel stores it via `_xlfn._LONGTEXT(...)`, which evaluates to
`#VALUE!`. So when a link URL exceeds 255 characters (e.g. some
Bluecore/Attentive `compose/design?...` URLs), `InstallTimedCampaignLink` skips
the timed formula and instead writes a real Excel hyperlink (its address has no
255-char limit) showing the clean platform name. Normal-length links keep the
live timed formula; long links stay clickable with no error.

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

The same highlighting is also applied to the **Dashboard** feed
(`DashboardWorkTable`, range `A11:L160`). Because the Dashboard feed has no
`Scheduled` column, "unscheduled" there is derived from the `Stage` column: a
feed row fills orange (Email) or yellow (SMS) when its `Send Date` is in the
same next-day/Friday-weekend window and its `Stage` is not yet `Scheduled` or
`Sent` (cancelled rows excluded). The Email rule is:

```
=AND($C11="Email",IFERROR(INT($A11),0)>=TODAY()+1,
  IFERROR(INT($A11),0)<=TODAY()+IF(WEEKDAY(TODAY(),1)=6,3,1),
  $F11<>"Scheduled",$F11<>"Sent",LEFT($D11,9)<>"CANCELLED")
```

The SMS rule uses `$C11="SMS"`. These rules are added below the existing
per-column Dashboard rules, so Approval/Segments/Stage indicators stay readable,
and they survive `RefreshDashboard`.

### Monthly Calendars

`June 2026 Calendar` (visible) and `May 2026 Calendar` (hidden) mirror the team's
SharePoint Email and SMS planning workbooks through external-link formulas such
as `='[3]2026 June Email Calendar'!$B$15`; there is no manual data entry. The
external links resolve to `adorama.sharepoint.com/.../EMAIL PLANNING/...` and
`/SMS Planning/...` paths and refresh in desktop Excel for a user with SharePoint
access (the QA harness allows `http(s)`/SharePoint links and disallows only local
`file://` links).

`Template for Duplicate` is a pre-formatted month used to add new calendars: copy
it, rename the copy to `<Month> 2026 Calendar`, then use **Data > Edit Links >
Change Source** to point it at that month's SharePoint Email and SMS files, and
update. Keep the calendar layout unchanged. The Template and Backup workbooks
include `June 2026 Calendar` and `Template for Duplicate` as a guided example;
the active tracker additionally keeps the hidden `May 2026 Calendar`.

`ValidateWorkbookConfiguration` no longer flags Calendar sheets (the retired-
calendar check was removed); it still rejects `DeliveredComparison` objects.

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

### Campaign Sheet Week Number Column

The `Email Campaigns` and `SMS Campaigns` tables each carry a `Week Number`
calculated column inserted as the **first** table column (column `A`), immediately
left of `Send Date` (which moves to column `B`). Unlike the Dashboard column, it
derives the week purely from the row's own `Send Date`:

```
=IF(ISNUMBER([@[Send Date]]),"Week "&WEEKNUM([@[Send Date]],1),"")
```

It shows a `Week N` label (Sunday-start `WEEKNUM(..,1)`) so campaigns can be grouped
and filtered by week for historical review; empty rows stay blank. Inserting the
column shifts the other columns right, and Excel automatically re-points the
existing conditional formatting (schedule-gap, Current Stage, Cancelled) and the
Dashboard's structured-reference formulas. VBA and QA both resolve campaign columns
by header name, so no further edits are needed and `ValidateWorkbookConfiguration`
stays `OK`. (Note: the legacy backup copy's campaign tables predate the `Scheduled`
column and therefore have one fewer column than the active tracker and template.)

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

`Notes - Instructions` is the in-workbook user guide. It is written in plain,
non-technical end-user language — a five-column table (*Feature*, *What you do*,
*How it works (in plain words)*, *Please do / Please avoid*, *Where it works*) —
and documents only features that are still in use; the title cell stays
`Detailed Notes and Instructions`. The same content is kept in sync across the
three workbooks. The embedded Notes-builder routines (`AddInstructionRow`) still
hold the older, more technical wording, so a full Notes rebuild would overwrite
the plain-language text — that rebuild is not part of normal use.

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
- Do not run the legacy migration or calendar-rebuild routines
  (`RebuildMonthlyCalendars`); the monthly calendars are SharePoint-linked, not
  VBA-built.

## Development and QA Tooling (`tools/`)

**Dependencies:** all scripts run on Python 3.14 in `.venv` and drive desktop
Excel through `win32com` (`pywin32`) COM automation, except the offline static
gate. Excel for the web cannot run them. None are required at workbook runtime.

### QA harnesses

| Script | Purpose | Behavior |
| --- | --- | --- |
| `qa_email_sms_campaign_tracker.py` | Primary non-destructive QA gate. | Copies the workbook to a temp file, opens it via COM, asserts VBA-continuation integrity, runs `ValidateWorkbookConfiguration`, then verifies KPI formulas, retired-comparison absence, audit fields, dropdowns, date/time display, timed hyperlinks, Notes password, and SharePoint-only external links. Prints `QA PASSED` + checklist. Default target: the active tracker. |
| `extensive_qa_email_sms_campaign_tracker.py` | Heavy scenario QA. | Seeds 10 Email + 10 SMS temporary campaigns in a disposable copy, exercises embedded VBA and native formulas end to end, then discards the copy. |
| `compile_vba_probe.py` | VBA project compile + modal-error detection. | Opens the workbook and forces a VBA compile, surfacing compiler dialogs as failures. |
| `inspect_workbook.py`, `quick_check.py`, `quick_test*.py` | Read-only spot checks. | Print sheet/table/formula state for manual inspection. |

An **offline static gate** (no Excel required) can pre-screen before the COM
harness: it un-doubles the double-spaced VBA exports and applies the same
blank-after-`_` continuation rule, and `py_compile`s every `tools/*.py`. Note the
loose dumps are snapshots, so this gate validates the exports' syntax, not the
live workbook — the COM harness above remains the authoritative release gate.

### Build, patch, and content scripts

| Script | Purpose |
| --- | --- |
| `build_clean_tracker.py`, `apply_modifications.py`, `sanitize_vba.py` | Assemble a clean workbook, apply structural modifications, and strip invalid VBA continuation gaps. |
| `export_vba.py` | Export the **current** VBA from the `.xlsm` (use this, not the stale dumps). |
| `fix_campaign_entry_formats.py` | Inject `ApplyCampaignEntryFormats` / `ApplyDashboardKpiFormulas` and resilient date/time behavior. |
| `split_campaign_type_and_others_prompt.py` | Split `Loyalty & PLCC` into independent options and inject `PromptCustomCampaignType`. |
| `add_week_number_source_column.py`, `add_dashboard_week_column.py` | Add the campaign-table and Dashboard `Week Number` columns. |
| `apply_schedule_highlight_and_weeknum.py`, `integrate_calendars_and_dashboard_highlight.py` | Add schedule-gap conditional formatting, the Week Number tile, and the SharePoint calendar integration. |
| `sync_calendars_to_template_backup.py`, `fix_backup_scheduled_column.py` | Propagate calendar examples to Template/Backup and bring the legacy backup's tables up to the current column set. |
| `rewrite_notes_for_end_users.py`, `update_instructions.py`, `remove_retired_notes_references.py` | Maintain the plain-language `Notes - Instructions` sheet. |
| `retire_calendar_comparisons.py` | One-time removal of the retired weekly delivery comparison blocks. |

## QA Release Process

A fast offline pre-check (un-doubled VBA continuation scan + `py_compile` of
`tools/`) can run first, but the authoritative gate is the COM harness. Each
release is checked for:

1. XLSM package and embedded VBA integrity.
2. VBA continuation syntax and full VBA project compilation.
3. Preservation of existing campaign data and formulas.
4. Ten temporary Email and ten temporary SMS campaign scenarios.
5. Checkbox, Current Stage, validation, filtering, date, time, and audit logic.
6. Dashboard date window, friendly statuses, cancellation exclusion, and KPIs.
7. Exact seven-day hyperlink behavior before and after the maturity timestamp.
8. Formula errors, broken names, freeze panes, and non-SharePoint external
   workbook links (SharePoint calendar links are allowed).
9. Instruction-sheet protection and password verification.
10. Schedule-gap highlighting rules and the Dashboard Week Number tile survive a
    Dashboard refresh, and the highlight rules add no formula errors.

Temporary QA records are created only in disposable copies.
