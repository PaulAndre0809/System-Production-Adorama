# Production Inventory Migration

Target workbook: `Email_Production_Inventory_Tracker_UI.xlsm`  
Target table: `ProductionInventoryTable` on `Production Inventory`

## Why the existing module must be replaced

The original VBA uses fixed column letters such as `A`, `F`, `V`, and `W`.
It also calculates stage and risk from columns that this migration removes.
Deleting columns without replacing that code would cause macros to write into
the wrong fields and would leave broken Dashboard references.

The replacement module keeps the existing public procedure names, but resolves
columns by table header. It also rebuilds the affected Dashboard formulas.

## Safe implementation

1. Close any other copies of the workbook and make sure it is saved as `.xlsm`.
2. Open the workbook and press `Alt+F11`.
3. In the Project Explorer, right-click `modEmailProductionTracker`, choose
   **Export File**, and retain that export as an additional code backup.
4. Right-click `modEmailProductionTracker` again and choose **Remove**. Choose
   **No** when Excel asks whether to export it again.
5. Choose **File > Import File** in the VBA editor and import
   `Email_Production_Inventory_Tracker_UI.bas`.
6. Choose **Debug > Compile VBAProject**. Resolve any compile error before
   continuing. A duplicate procedure error means the old module was not removed.
7. Return to Excel, press `Alt+F8`, select
   `MigrateProductionInventoryStructure`, and click **Run**.
8. The macro creates a timestamped `*_PRE_MIGRATION_*.xlsm` backup before making
   changes. Review the final table, Dashboard, formulas, and any UserForms.
9. Save the reviewed workbook as `.xlsm`.

## Resulting Production Inventory order

The eight checklist columns are inserted immediately after `Owner`:

1. `Send Date`
2. `Send Time`
3. `Campaign Name`
4. `Campaign Type`
5. `Current Stage`
6. `Owner`
7. `Campaign Name and UTM Parameter (Source Code)`
8. `Creative Brief, SL & PH`
9. `SKUs`
10. `In-Design`
11. `Build, QA`
12. `Route`
13. `Approval`
14. `Segments`
15. `Jira Link`
16. `ClickUp Link`
17. `Bluecore Link`
18. `Est. Audience`
19. `Delivered`
20. `Last Updated`

Each checklist column receives a data-validation dropdown containing unchecked
and checked symbols. `Send Date` receives the `MM/DD/YYYY` number format.

## UserForm compatibility check

The provided module preserves known public macro names. If a separate UserForm
directly uses expressions such as `Cells(row, "V")` or `Range("H:H")`, update it
to call `InventoryColumnNumber("Header Name")`. Direct hard-coded references in
code outside the exported module cannot be repaired automatically without
reviewing that code.

If the migration reports new `#REF!` references, do not save the changed
workbook. Close it and restore the timestamped pre-migration backup.

## Monthly calendars

The consolidated `Email_Production_Inventory_Tracker_UI.bas` module includes
the `RebuildMonthlyCalendars` maintenance macro. It creates January through
December calendar sheets for the current year. Calendar cells read `Send Date`
and `Campaign Name` directly from `ProductionInventoryTable`, so changes to the
inventory appear automatically.

The calendar rebuild also:

- Adds January-December navigation links to the Dashboard.
- Hides the required `Dropdowns` and `Automation Log` support sheets.
- Removes the obsolete `README` and `VBA Code` worksheets.
- Applies freeze panes, print areas, gridline settings, and consistent styling.
