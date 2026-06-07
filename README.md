<div align="center">

# 📊 System Production Adorama

[![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)](Production%20Tracker/)
[![Status](https://img.shields.io/badge/Status-Active-success?style=for-the-badge)](#)
[![Tracker](https://img.shields.io/badge/Version-v2.0-blue?style=for-the-badge)](#)

*Production tracking tools and systems, primarily focusing on email and SMS campaign management and reporting analysis.*

</div>

---

## 🚀 Components

### 📧 Email & SMS Campaign Tracker

Located in the [`Production Tracker`](Production%20Tracker/) directory, this tool is an advanced Excel-based application (`.xlsm`) designed for tracking the workflow, stages, and schedules of Email and SMS marketing campaigns.

<details open>
<summary><b>📂 Worksheet Structure & Navigation</b></summary>
<br>

The workbook is fully self-contained and structured logically across the following worksheets:

| Sheet Name | Description |
| :--- | :--- |
| 📝 **`Notes - Instruction`** | Comprehensive user and technical guide detailing the schema rules, stability requirements, and instructions for SharePoint collaboration. |
| 🎛️ **`Dashboard`** | Combined command center displaying active work, open campaigns, campaigns scheduled for the current and next week, and a delivered comparison table. |
| 📧 **`Email Campaigns`** | Source table tracking the email workflow, utilizing checkbox progression to drive campaign status. |
| 📱 **`SMS Campaigns`** | Source table tracking the SMS workflow, mirroring the email metadata schema but with a streamlined checklist. |
| 📅 **`Monthly Calendars`** | Dynamic calendars (January to December) aggregating both email and SMS campaigns sorted by `Send Date`. The current month tab is dynamically highlighted in green. |
| 🔽 **`Dropdowns`** | Storage for valid metadata values (e.g., Campaign Types, Owner list). |
| 📜 **`Automation Log`** | Running ledger of system updates and audit logs. |

</details>

<br>

<details>
<summary><b>📊 Data Schemas & Table Headers</b></summary>
<br>

#### 📧 Email Campaigns Table
Tracks the full email deployment pipeline with **8 checkbox columns** (G:N):

*   **Core Metadata**: `Send Date`, `Send Time`, `Campaign Name`, `Campaign Type`, `Current Stage`, `Owner`
*   **Workflow Checkboxes**:
    1. `Campaign Name and UTM Parameter (Source Code)`
    2. `Creative Brief, SL & PH`
    3. `SKUs`
    4. `In-Design`
    5. `Build, QA`
    6. `Route`
    7. `Approval`
    8. `Segments`
*   **Links & Audit**: `Jira Link`, `ClickUp Link`, `Bluecore/Attentive Link`, `Est. Audience`, `Delivered`, `Last Updated`, `Last Updated By`, `Notes`

#### 📱 SMS Campaigns Table
Streamlines the SMS pipeline with **4 checkbox columns** (G:J):

*   **Core Metadata**: `Send Date`, `Send Time`, `Campaign Name`, `Campaign Type`, `Current Stage`, `Owner`
*   **Workflow Checkboxes**:
    1. `Send SMS Options`
    2. `Send Test`
    3. `Approval`
    4. `Segments`
*   **Links & Audit**: `Jira Link`, `ClickUp Link`, `Bluecore/Attentive Link`, `Est. Audience`, `Delivered`, `Last Updated`, `Last Updated By`, `Notes`

</details>

<br>

<details>
<summary><b>⚙️ State Machine & Stage Calculation</b></summary>
<br>

The `Current Stage` is an automatic, formula-driven column on both campaign sheets:

> **📧 Email Workflow Stages**
> `No checklist items checked` ➔ `Source Code` ➔ `Creative Brief` ➔ `Waiting for SKUs` ➔ `With Design` ➔ `Build / QA` ➔ `Routing` ➔ `Awaiting Approval` ➔ `Segments` ➔ `Links Pending` ➔ `Ready to Schedule` ➔ `Scheduled` ➔ `Sent` *(when `Delivered` > 0)*.

> **📱 SMS Workflow Stages**
> `SMS Options` ➔ `Send Test` ➔ `Awaiting Approval` ➔ `Segments` ➔ `Links Pending` ➔ `Ready to Schedule` ➔ `Scheduled` ➔ `Sent` *(when `Delivered` > 0)*.

</details>

<br>

<details>
<summary><b>☁️ Compatibility & SharePoint Coauthoring</b></summary>
<br>

*   ✅ **Native Checkboxes**: Built using modern Microsoft 365 native in-cell checkboxes (evaluating to true Boolean `TRUE`/`FALSE`). Legacy systems double-click cells for a visual Boolean fallback.
*   🌐 **Excel for the Web**: Supports web-based edits, coauthoring, and automated formula calculations.
*   💻 **VBA Macros**: Microsoft limits VBA execution to Desktop Excel. Consequently, automatic updater auditing (timestamps/editor user logging), dashboard snapshot generation, and daily digests require opening the workbook in Desktop Excel.

> [!WARNING]
> **Macro Limitations on the Web:** Ensure you open the workbook in the Desktop App if you need to run automated reporting tools or use audit stamping!

</details>

---

## 📈 Reporting Analysis

Located in the [`Reporting Analysis`](Reporting%20Analysis/) directory. *(Currently empty / under development)*.

<br>

<div align="center">
  <sub>Built with ❤️ for Production Teams</sub>
</div>
