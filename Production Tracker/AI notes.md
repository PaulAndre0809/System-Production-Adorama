# 🤖 AI Development Notes - Production Tracker

This document establishes development standards, architecture notes, update protocols, and consistency guides for AI assistants collaborating on the **System Production Adorama** tracker codebase.

---

## 📂 System Architecture Overview

The Production Tracker is an Excel Macro-Enabled Workbook (`.xlsm`) designed to manage the rollout of outreach campaigns. It is built to support coauthoring via SharePoint (Excel for the Web) while retaining advanced local automation via VBA.

### Worksheet Architecture
| Sheet Name | Type | Key Purpose / Data Contents |
| :--- | :--- | :--- |
| **`Notes - Instruction`** | Text Info | Interactive user and technical guide outlining table layouts, checklist rules, and stability requirements. |
| **`Dashboard`** | Dynamic View | Combined command center showing active work, today's sends, weekly metrics, and a dynamic view of current + next week campaigns. |
| **`Email Campaigns`** | Data Table | Source table tracking the email workflow, using checkbox progression to drive campaign status (`EmailCampaignsTable`). |
| **`SMS Campaigns`** | Data Table | Source table tracking the SMS workflow, mirroring the email metadata schema but with a streamlined checklist (`SMSCampaignsTable`). |
| **`Monthly Calendars`** | Dynamic View | Rebuilt monthly tabs (Jan - Dec) showing both email and SMS campaigns sorted by `Send Date`. |
| **`Dropdowns`** | Data Source | Stores validation list options (Campaign Types, Owners). Very hidden sheet. |
| **`Automation Log`** | Data Table | Running ledger of automated operations, timestamping, and editor user logging. Very hidden sheet. |

### Automation Architecture
- **Web App Recalculations**: Simple formulas, the Dashboard `VSTACK`/`FILTER` spill formulas (`$AA$11#`), calendar bullet joins, and charts are calculated natively in the web browser.
- **VBA Modules**: Formats hyperlinks, structures tables, builds calendars dynamically, and applies daily digests. VBA runs **only in Desktop Excel**.
- **Audit System**: Automatically stamps `Last Updated` and `Last Updated By` when users edit rows locally.

---

## 📅 Dynamic Highlighting & Filtering Logic

To assist production teams in identifying high-priority deliverables, upcoming campaigns are dynamically filtered and visually highlighted on the **`Dashboard`** worksheet within the **`DashboardWorkTable`** queue.

### Temporal Logic Rules
1. **Weekday Schedule (Monday - Thursday)**:
   - Target Date: **Tomorrow** (`Today + 1`).
   - Action: Filters the `DashboardWorkTable` queue to show only tomorrow's rollouts and highlights their rows in a soft theme-compatible color.
2. **Weekend & Monday Schedule (Friday)**:
   - Target Dates: **Saturday, Sunday, and Monday** (`Today + 1` to `Today + 3`).
   - Action: Filters the `DashboardWorkTable` queue to show the weekend and Monday rollout queue and highlights their rows.
3. **Weekend Fallback (Saturday - Sunday)**:
   - Target Date: **Tomorrow** (`Today + 1`).
   - Action: Fallback highlight and filter for tomorrow.

### Implementation Standards (VBA)
- Use **`CDbl(targetDate)`** when passing criteria to `AutoFilter`. VBA `AutoFilter` is notoriously locale-sensitive with date strings; using the double-precision numeric serial representation bypasses date-format mismatch bugs entirely.
- Apply a **Premium Theme Accent** instead of standard Excel colors (e.g., use a soft blue `RGB(221, 235, 247)` to match the workbook style).
- Ensure filtering and highlighting can be cleared cleanly using a `ResetDashboardFilters` subroutine to restore standard table visibility.
- Integrate the function into the `RefreshDashboard` process so the highlighting updates automatically.

---

## 💻 Codebase Formatting & Standards

When writing or updating VBA scripts or Excel formula patterns, strictly adhere to the following standards:

### 1. No Hardcoded Column Letters
Never reference column letters (e.g., `Range("G:N")` or `Cells(r, 7)`) for data storage or checkbox calculations. The schema can change, which will break hardcoded column structures.
- **Rule**: Dynamically lookup column indices by header name.
- **Approved Helpers**:
  - `FindTableColumn(lo, headerName)`: Finds the matching `ListColumn` in a table.
  - `ValueByHeader(ws, rowNumber, headerName)`: Retrieves a row cell value by header title.

### 2. Native Checkbox Usage
- Modern columns use native Microsoft 365 in-cell checkboxes (evaluating to true Boolean `TRUE`/`FALSE`).
- Legacy fallback sheets use segui UI symbols (`ChrW(&H2611)` / `ChrW(&H2610)`) with custom number formatting (`[=1]"\u2611";[=0]"\u2610"`). Always verify the control type using `CheckboxControlType(rng)` before setting values.

### 3. Calculation & Performance Optimization
VBA routines operating on large ranges must temporarily toggle Excel application settings to prevent screen-flickering and calculation lag:
```vba
Dim oldCalculation As XlCalculation
oldCalculation = Application.Calculation
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
Application.EnableEvents = False

' ... perform operations ...

Application.Calculation = oldCalculation
Application.ScreenUpdating = True
Application.EnableEvents = True
```

### 4. Error Handling
All public VBA procedures must implement basic error handlers to restore application settings (events, calculations) and alert the user cleanly:
```vba
On Error GoTo ErrorHandler
' ... code ...
CleanExit:
    ' Restore settings
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume CleanExit
```

---

## 🤖 Guidelines for AI Assistants

To maintain codebase safety and avoid disrupting users, review these constraints prior to making modifications:

- **Preserve Stage Formulas**: The `Current Stage` column uses an automatic formula that concatenates checked checkboxes. Do not overwrite this with static values.
- **Maintain Tab Color Rules**: Tab colors are styled dynamically (e.g., current month calendar tab is green `RGB(0, 176, 80)`). Ensure these colors are left alone or computed dynamically using the sheet-style routines.
- **Table Integrity**: Check for `#REF!` errors in formulas and workbook names. Run `ValidateWorkbookConfiguration()` (or its equivalence) to verify configuration validity.
- **Activity Logging**: Any script-driven configuration change or bulk action should write a row to the `Automation Log` using the `LogAction(actionName, details)` helper.
