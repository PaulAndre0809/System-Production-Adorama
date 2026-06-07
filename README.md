# System Production Adorama

This repository contains production tracking tools and systems, primarily focusing on email and SMS campaign management and reporting analysis.

---

## Components

### Email & SMS Campaign Tracker

Located in the [`Production Tracker`](Production%20Tracker/) directory, this tool is an Excel-based application (`.xlsm`) designed for tracking the workflow, stages, and schedules of Email and SMS marketing campaigns.

#### Worksheet Structure & Navigation

The workbook is fully self-contained and structured logically across the following worksheets:

*   **`Notes - Instruction`**: Comprehensive user and technical guide detailing the schema rules, stability requirements, and instructions for SharePoint collaboration.
*   **`Dashboard`**: Combined command center displaying active work, open campaigns, campaigns scheduled for the current and next week (Sunday through Saturday), and a delivered comparison table.
*   **`Email Campaigns`**: Source table tracking the email workflow, utilizing checkbox progression to drive campaign status.
*   **`SMS Campaigns`**: Source table tracking the SMS workflow, mirroring the email metadata schema but with a streamlined checklist.
*   **Monthly Calendars (`January Calendar` to `December Calendar`)**: Dynamic calendars aggregating both email and SMS campaigns sorted by `Send Date`. The current month tab is dynamically highlighted in green.
*   **`Dropdowns`**: Storage for valid metadata values (e.g., Campaign Types, Owner list).
*   **`Automation Log`**: Running ledger of system updates and audit logs.

---

#### Data Schemas & Table Headers

##### Email Campaigns Table
Tracks the full email deployment pipeline with **8 checkbox columns** (G:N):
*   **Core Metadata**: `Send Date`, `Send Time`, `Campaign Name`, `Campaign Type`, `Current Stage`, `Owner`
*   **Workflow Checkboxes**:
    1.  `Campaign Name and UTM Parameter (Source Code)`
    2.  `Creative Brief, SL & PH`
    3.  `SKUs`
    4.  `In-Design`
    5.  `Build, QA`
    6.  `Route`
    7.  `Approval`
    8.  `Segments`
*   **Links & Audit**: `Jira Link`, `ClickUp Link`, `Bluecore/Attentive Link`, `Est. Audience`, `Delivered`, `Last Updated`, `Last Updated By`, `Notes`

##### SMS Campaigns Table
Streamlines the SMS pipeline with **4 checkbox columns** (G:J):
*   **Core Metadata**: `Send Date`, `Send Time`, `Campaign Name`, `Campaign Type`, `Current Stage`, `Owner`
*   **Workflow Checkboxes**:
    1.  `Send SMS Options`
    2.  `Send Test`
    3.  `Approval`
    4.  `Segments`
*   **Links & Audit**: `Jira Link`, `ClickUp Link`, `Bluecore/Attentive Link`, `Est. Audience`, `Delivered`, `Last Updated`, `Last Updated By`, `Notes`

---

#### State Machine & Stage Calculation

The `Current Stage` is an automatic, formula-driven column on both campaign sheets:

*   **Email Workflow Stages**: Progresses sequentially through `No checklist items checked` $\rightarrow$ `Source Code` $\rightarrow$ `Creative Brief` $\rightarrow$ `Waiting for SKUs` $\rightarrow$ `With Design` $\rightarrow$ `Build / QA` $\rightarrow$ `Routing` $\rightarrow$ `Awaiting Approval` $\rightarrow$ `Segments` $\rightarrow$ `Links Pending` $\rightarrow$ `Ready to Schedule` $\rightarrow$ `Scheduled` $\rightarrow$ `Sent` (when `Delivered` $> 0$).
*   **SMS Workflow Stages**: Progresses sequentially through `SMS Options` $\rightarrow$ `Send Test` $\rightarrow$ `Awaiting Approval` $\rightarrow$ `Segments` $\rightarrow$ `Links Pending` $\rightarrow$ `Ready to Schedule` $\rightarrow$ `Scheduled` $\rightarrow$ `Sent` (when `Delivered` $> 0$).

---

#### Compatibility & SharePoint Coauthoring

*   **Native Checkboxes**: Built using modern Microsoft 365 native in-cell checkboxes (evaluating to true Boolean `TRUE`/`FALSE`). Legacy systems double-click cells for a visual Boolean fallback.
*   **Excel for the Web**: Supports web-based edits, coauthoring, and automated formula calculations.
*   **VBA Macros**: Microsoft limits VBA execution to Desktop Excel. Consequently, automatic updater auditing (timestamps/editor user logging), dashboard snapshot generation, and daily digests require opening the workbook in Desktop Excel.

---

### Reporting Analysis

Located in the [`Reporting Analysis`](Reporting%20Analysis/) directory. (Currently empty / under development).

---

## Setup & Maintenance

The repository includes a Python virtual environment structure (`.venv`) intended for running the provided QA and maintenance scripts.

### Transactional Deployment

Close all Excel windows and run the PowerShell deployer to update workbook macros and schemas safely without risking file corruption:

```powershell
powershell -ExecutionPolicy Bypass -File ".\Production Tracker\Email & SMS Campaign Tracker.ps1"
```

The script runs a safe, multi-step deployment pipeline:
1.  Creates a temporary copy of the tracker.
2.  Sets calculation to manual and purges legacy VBA components.
3.  Injects updated VBA modules (`modEmailProductionTracker`, `ThisWorkbook` handlers, and sheet-level event macros).
4.  Compiles the VBA code to prevent runtime execution errors.
5.  Applies workbook-wide configurations, style layouts, and schemas.
6.  Saves, tests, and validates the copy.
7.  Swaps the production tracker file only if validation succeeds, rolling back immediately on failure.

### Quality Assurance & Validation Suite

Run the suite to perform comprehensive validation across schemas, behavior, and persistence:

```powershell
.venv\Scripts\python.exe ".\Production Tracker\Email & SMS Campaign Tracker.py" --qa
```

The validation suite verifies:
*   **Structural integrity**: Table names, column positions, header labels.
*   **Checkbox functionality**: Interaction behavior and Boolean evaluation.
*   **Stage transitions**: Correct formula calculation for both Email and SMS campaign workflows.
*   **Audit logs**: Correct updater username and timestamp generation.
*   **Aggregation**: Proper filtering and display on Dashboard and Month calendars.
*   **Performance metrics**: Checks file-save times, workbook recalculation speed, and confirms automatic calculation is active.

