# System Production Adorama

This repository contains production tracking tools and systems, primarily focusing on email and SMS campaign management and reporting analysis.

## Components

### Email & SMS Campaign Tracker
Located in the [`Production Tracker`](Production%20Tracker/) directory, this tool is an Excel-based application (`.xlsm`) designed for tracking the workflow and progress of Email and SMS marketing campaigns.

**Key features:**
- **Campaign Sheets:** Dedicated tracking sheets for Email and SMS workflows with native checkbox progression.
- **Dynamic Stages:** Automated calculation of campaign stages based on task completion.
- **Dashboards & Calendars:** A combined visual dashboard for a 14-day schedule snapshot and comprehensive monthly calendar aggregations.
- **Reporting:** Built-in comparisons for delivered campaigns (e.g., Week-over-Week).
- **VBA & Automation:** Embedded VBA macros for updater auditing, dashboard generation, and daily digests.
- **QA & Deployment Utilities:** PowerShell (`.ps1`) and Python (`.py`) scripts for automated, transactional deployment and quality assurance testing.

For detailed instructions and technical documentation, refer to the [Email & SMS Campaign Tracker Guide](Production%20Tracker/Email%20&%20SMS%20Campaign%20Tracker.md).

### Reporting Analysis
Located in the `Reporting Analysis` directory. (Currently empty / under development).

## Setup & Maintenance

The repository includes a Python virtual environment structure (`.venv`) intended for running the provided QA and maintenance scripts.

### Deployment & QA

Deployment of the Excel tracker can be executed transactionally using the provided PowerShell script to ensure data integrity and schema validation:
```powershell
powershell -ExecutionPolicy Bypass -File ".\Production Tracker\Email & SMS Campaign Tracker.ps1"
```

A Python suite is available for rigorous quality assurance and maintenance testing:
```powershell
.venv\Scripts\python.exe ".\Production Tracker\Email & SMS Campaign Tracker.py" --qa
```
