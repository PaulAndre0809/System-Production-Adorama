# Systemic Operational Mechanics

## Data Flow
- Campaigns are entered in Email Campaigns or SMS Campaigns tables.
- Dashboard is updated via RefreshDashboard without rebuilding source tables.
- Checkboxes and formulas manage state.

## Control Flow
- Event modules (Worksheet_Change) trigger HandleCampaignChange which orchestrates formatting, timestamps, and Dashboard updates.

## Key Dependencies
- win32com for python automation.
- SharePoint for monthly calendars.

## High-Level Architecture
- VBA standard module modEmailProductionTracker acts as the core engine. Event modules are thin delegators.


# Module / File: ThisWorkbook

## Function: Workbook_Open
- **Purpose**: Entry point for Workbook_Open.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for Workbook_Open.
- **Side Effects**: Modifies workbook state.

## Function: Workbook_BeforeClose
- **Purpose**: Entry point for Workbook_BeforeClose.
- **Inputs**:
  - Cancel (Boolean): Parameter Cancel
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for Workbook_BeforeClose.
- **Side Effects**: Modifies workbook state.

# Module / File: Sheet2

# Module / File: Sheet1

## Function: Worksheet_BeforeDoubleClick
- **Purpose**: Entry point for Worksheet_BeforeDoubleClick.
- **Inputs**:
  - Target (Range): Parameter Target
  - Cancel (Boolean): Parameter Cancel
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for Worksheet_BeforeDoubleClick.
- **Side Effects**: Modifies workbook state.

## Function: Worksheet_Change
- **Purpose**: Entry point for Worksheet_Change.
- **Inputs**:
  - Target (Range): Parameter Target
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for Worksheet_Change.
- **Side Effects**: Modifies workbook state.

# Module / File: Sheet4

# Module / File: Sheet5

# Module / File: Sheet17

## Function: Worksheet_BeforeDoubleClick
- **Purpose**: Entry point for Worksheet_BeforeDoubleClick.
- **Inputs**:
  - Target (Range): Parameter Target
  - Cancel (Boolean): Parameter Cancel
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for Worksheet_BeforeDoubleClick.
- **Side Effects**: Modifies workbook state.

## Function: Worksheet_Change
- **Purpose**: Entry point for Worksheet_Change.
- **Inputs**:
  - Target (Range): Parameter Target
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for Worksheet_Change.
- **Side Effects**: Modifies workbook state.

# Module / File: modEmailProductionTracker

## Function: MigrateProductionInventoryStructure
- **Purpose**: Entry point for MigrateProductionInventoryStructure.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for MigrateProductionInventoryStructure.
- **Side Effects**: Modifies workbook state.

## Function: RefreshProductionStatus
- **Purpose**: Entry point for RefreshProductionStatus.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for RefreshProductionStatus.
- **Side Effects**: Modifies workbook state.

## Function: ApplyAllConfigurations
- **Purpose**: Entry point for ApplyAllConfigurations.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ApplyAllConfigurations.
- **Side Effects**: Modifies workbook state.

## Function: ApplyDashboardKpiFormulas
- **Purpose**: Entry point for ApplyDashboardKpiFormulas.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ApplyDashboardKpiFormulas.
- **Side Effects**: Modifies workbook state.

## Function: ApplyTimedCampaignLinks
- **Purpose**: Entry point for ApplyTimedCampaignLinks.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ApplyTimedCampaignLinks.
- **Side Effects**: Modifies workbook state.

## Function: CampaignLinkAddress
- **Purpose**: Calculates or retrieves CampaignLinkAddress.
- **Inputs**:
  - cell (Range): Parameter cell
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CampaignLinkAddress.
- **Side Effects**: None

## Function: RefreshTimedCampaignLinks
- **Purpose**: Entry point for RefreshTimedCampaignLinks.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for RefreshTimedCampaignLinks.
- **Side Effects**: Modifies workbook state.

## Function: CalculateTimedCampaignLinkColumns
- **Purpose**: Entry point for CalculateTimedCampaignLinkColumns.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CalculateTimedCampaignLinkColumns.
- **Side Effects**: Modifies workbook state.

## Function: ScheduleNextCampaignLinkRefresh
- **Purpose**: Entry point for ScheduleNextCampaignLinkRefresh.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ScheduleNextCampaignLinkRefresh.
- **Side Effects**: Modifies workbook state.

## Function: CancelCampaignLinkRefresh
- **Purpose**: Entry point for CancelCampaignLinkRefresh.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CancelCampaignLinkRefresh.
- **Side Effects**: Modifies workbook state.

## Function: NextCampaignLinkMaturity
- **Purpose**: Calculates or retrieves NextCampaignLinkMaturity.
- **Inputs**:
  - None
- **Outputs**: Date
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for NextCampaignLinkMaturity.
- **Side Effects**: None

## Function: RefreshDashboard
- **Purpose**: Entry point for RefreshDashboard.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for RefreshDashboard.
- **Side Effects**: Modifies workbook state.

## Function: RefreshNativeOutputs
- **Purpose**: Entry point for RefreshNativeOutputs.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for RefreshNativeOutputs.
- **Side Effects**: Modifies workbook state.

## Function: ApplyDashboardAuditHeader
- **Purpose**: Entry point for ApplyDashboardAuditHeader.
- **Inputs**:
  - ws (Worksheet): Parameter ws
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ApplyDashboardAuditHeader.
- **Side Effects**: Modifies workbook state.

## Function: DashboardLastEditedByFormula
- **Purpose**: Calculates or retrieves DashboardLastEditedByFormula.
- **Inputs**:
  - None
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for DashboardLastEditedByFormula.
- **Side Effects**: None

## Function: DashboardNativeSpillFormula
- **Purpose**: Calculates or retrieves DashboardNativeSpillFormula.
- **Inputs**:
  - None
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for DashboardNativeSpillFormula.
- **Side Effects**: None

## Function: DashboardNativeBlankRowFormula
- **Purpose**: Calculates or retrieves DashboardNativeBlankRowFormula.
- **Inputs**:
  - None
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for DashboardNativeBlankRowFormula.
- **Side Effects**: None

## Function: ApprovalDashboardStatus
- **Purpose**: Calculates or retrieves ApprovalDashboardStatus.
- **Inputs**:
  - statusValue (Variant): Parameter statusValue
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ApprovalDashboardStatus.
- **Side Effects**: None

## Function: SegmentsDashboardStatus
- **Purpose**: Calculates or retrieves SegmentsDashboardStatus.
- **Inputs**:
  - statusValue (Variant): Parameter statusValue
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for SegmentsDashboardStatus.
- **Side Effects**: None

## Function: RemoveLegacyCalendarAndComparisonArtifacts
- **Purpose**: Entry point for RemoveLegacyCalendarAndComparisonArtifacts.
- **Inputs**:
  - wb (Workbook): Parameter wb
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for RemoveLegacyCalendarAndComparisonArtifacts.
- **Side Effects**: Modifies workbook state.

## Function: RemoveLegacyDashboardComparisons
- **Purpose**: Entry point for RemoveLegacyDashboardComparisons.
- **Inputs**:
  - ws (Worksheet): Parameter ws
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for RemoveLegacyDashboardComparisons.
- **Side Effects**: Modifies workbook state.

## Function: IsLegacyCalendarSheetName
- **Purpose**: Calculates or retrieves IsLegacyCalendarSheetName.
- **Inputs**:
  - sheetName (String): Parameter sheetName
- **Outputs**: Boolean
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for IsLegacyCalendarSheetName.
- **Side Effects**: None

## Function: ApplyDashboardStatusFormatting
- **Purpose**: Entry point for ApplyDashboardStatusFormatting.
- **Inputs**:
  - dashboardTable (ListObject): Parameter dashboardTable
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ApplyDashboardStatusFormatting.
- **Side Effects**: Modifies workbook state.

## Function: CreateDailyDigest
- **Purpose**: Entry point for CreateDailyDigest.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CreateDailyDigest.
- **Side Effects**: Modifies workbook state.

## Function: LogAction
- **Purpose**: Entry point for LogAction.
- **Inputs**:
  - actionName (String): Parameter actionName
  - details (String): Parameter details
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for LogAction.
- **Side Effects**: Modifies workbook state.

## Function: InventoryColumnNumber
- **Purpose**: Calculates or retrieves InventoryColumnNumber.
- **Inputs**:
  - headerName (String): Parameter headerName
- **Outputs**: Long
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for InventoryColumnNumber.
- **Side Effects**: None

## Function: ApplyNativeCheckboxControl
- **Purpose**: Calculates or retrieves ApplyNativeCheckboxControl.
- **Inputs**:
  - rng (Range): Parameter rng
- **Outputs**: Boolean
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ApplyNativeCheckboxControl.
- **Side Effects**: None

## Function: CheckboxControlType
- **Purpose**: Calculates or retrieves CheckboxControlType.
- **Inputs**:
  - rng (Range): Parameter rng
- **Outputs**: Long
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CheckboxControlType.
- **Side Effects**: None

## Function: ApplyLegacyCheckboxDisplay
- **Purpose**: Entry point for ApplyLegacyCheckboxDisplay.
- **Inputs**:
  - rng (Range): Parameter rng
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ApplyLegacyCheckboxDisplay.
- **Side Effects**: Modifies workbook state.

## Function: SetChecklistValue
- **Purpose**: Entry point for SetChecklistValue.
- **Inputs**:
  - targetCell (Range): Parameter targetCell
  - isComplete (Boolean): Parameter isComplete
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for SetChecklistValue.
- **Side Effects**: Modifies workbook state.

## Function: LegacyCheckboxNumberFormat
- **Purpose**: Calculates or retrieves LegacyCheckboxNumberFormat.
- **Inputs**:
  - None
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for LegacyCheckboxNumberFormat.
- **Side Effects**: None

## Function: EnsureNotesColumn
- **Purpose**: Entry point for EnsureNotesColumn.
- **Inputs**:
  - lo (ListObject): Parameter lo
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for EnsureNotesColumn.
- **Side Effects**: Modifies workbook state.

## Function: ConfigureOwnerColumn
- **Purpose**: Entry point for ConfigureOwnerColumn.
- **Inputs**:
  - lo (ListObject): Parameter lo
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ConfigureOwnerColumn.
- **Side Effects**: Modifies workbook state.

## Function: ConfigureCampaignTypeColumn
- **Purpose**: Entry point for ConfigureCampaignTypeColumn.
- **Inputs**:
  - lo (ListObject): Parameter lo
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ConfigureCampaignTypeColumn.
- **Side Effects**: Modifies workbook state.

## Function: ConfigureSmsCampaignTypeColumn
- **Purpose**: Entry point for ConfigureSmsCampaignTypeColumn.
- **Inputs**:
  - lo (ListObject): Parameter lo
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ConfigureSmsCampaignTypeColumn.
- **Side Effects**: Modifies workbook state.

## Function: CampaignTypeOptions
- **Purpose**: Calculates or retrieves CampaignTypeOptions.
- **Inputs**:
  - None
- **Outputs**: Variant
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CampaignTypeOptions.
- **Side Effects**: None

## Function: CampaignTypeOptionsRange
- **Purpose**: Calculates or retrieves CampaignTypeOptionsRange.
- **Inputs**:
  - None
- **Outputs**: Range
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CampaignTypeOptionsRange.
- **Side Effects**: None

## Function: EnsureUpdatedByColumn
- **Purpose**: Entry point for EnsureUpdatedByColumn.
- **Inputs**:
  - lo (ListObject): Parameter lo
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for EnsureUpdatedByColumn.
- **Side Effects**: Modifies workbook state.

## Function: HandleInventorySelection
- **Purpose**: Entry point for HandleInventorySelection.
- **Inputs**:
  - Target (Range): Parameter Target
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for HandleInventorySelection.
- **Side Effects**: Modifies workbook state.

## Function: ToggleInventoryChecklist
- **Purpose**: Calculates or retrieves ToggleInventoryChecklist.
- **Inputs**:
  - Target (Range): Parameter Target
- **Outputs**: Boolean
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ToggleInventoryChecklist.
- **Side Effects**: None

## Function: HandleInventoryChange
- **Purpose**: Entry point for HandleInventoryChange.
- **Inputs**:
  - Target (Range): Parameter Target
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for HandleInventoryChange.
- **Side Effects**: Modifies workbook state.

## Function: CalculateStageRow
- **Purpose**: Entry point for CalculateStageRow.
- **Inputs**:
  - rowNumber (Long): Parameter rowNumber
  - lo (ListObject): Parameter lo
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CalculateStageRow.
- **Side Effects**: Modifies workbook state.

## Function: UpdateRowTimestampAndUser
- **Purpose**: Entry point for UpdateRowTimestampAndUser.
- **Inputs**:
  - rowNumber (Long): Parameter rowNumber
  - lo (ListObject): Parameter lo
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for UpdateRowTimestampAndUser.
- **Side Effects**: Modifies workbook state.

## Function: CurrentUserName
- **Purpose**: Calculates or retrieves CurrentUserName.
- **Inputs**:
  - None
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CurrentUserName.
- **Side Effects**: None

## Function: UpdateCalendarTabs
- **Purpose**: Entry point for UpdateCalendarTabs.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for UpdateCalendarTabs.
- **Side Effects**: Modifies workbook state.

## Function: ApplyCampaignEntryFormats
- **Purpose**: Entry point for ApplyCampaignEntryFormats.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ApplyCampaignEntryFormats.
- **Side Effects**: Modifies workbook state.

## Function: FormatSendDateColumn
- **Purpose**: Entry point for FormatSendDateColumn.
- **Inputs**:
  - lo (ListObject): Parameter lo
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for FormatSendDateColumn.
- **Side Effects**: Modifies workbook state.

## Function: FormatSendTimeColumn
- **Purpose**: Entry point for FormatSendTimeColumn.
- **Inputs**:
  - lo (ListObject): Parameter lo
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for FormatSendTimeColumn.
- **Side Effects**: Modifies workbook state.

## Function: ApplyCalculatedColumns
- **Purpose**: Entry point for ApplyCalculatedColumns.
- **Inputs**:
  - lo (ListObject): Parameter lo
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ApplyCalculatedColumns.
- **Side Effects**: Modifies workbook state.

## Function: CurrentStageFormulaText
- **Purpose**: Calculates or retrieves CurrentStageFormulaText.
- **Inputs**:
  - isSmsTable (Boolean): Parameter isSmsTable
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CurrentStageFormulaText.
- **Side Effects**: None

## Function: ChecklistStatusFormulaText
- **Purpose**: Calculates or retrieves ChecklistStatusFormulaText.
- **Inputs**:
  - isSmsTable (Boolean): Parameter isSmsTable
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ChecklistStatusFormulaText.
- **Side Effects**: None

## Function: FormulaString
- **Purpose**: Calculates or retrieves FormulaString.
- **Inputs**:
  - TextValue (String): Parameter TextValue
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for FormulaString.
- **Side Effects**: None

## Function: FormatLastUpdatedColumn
- **Purpose**: Entry point for FormatLastUpdatedColumn.
- **Inputs**:
  - updatedColumn (ListColumn): Parameter updatedColumn
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for FormatLastUpdatedColumn.
- **Side Effects**: Modifies workbook state.

## Function: EnsureCampaignSheets
- **Purpose**: Entry point for EnsureCampaignSheets.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for EnsureCampaignSheets.
- **Side Effects**: Modifies workbook state.

## Function: GetInventoryTable
- **Purpose**: Calculates or retrieves GetInventoryTable.
- **Inputs**:
  - None
- **Outputs**: ListObject
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for GetInventoryTable.
- **Side Effects**: None

## Function: GetSmsTable
- **Purpose**: Calculates or retrieves GetSmsTable.
- **Inputs**:
  - None
- **Outputs**: ListObject
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for GetSmsTable.
- **Side Effects**: None

## Function: NormalizeHeaderKey
- **Purpose**: Calculates or retrieves NormalizeHeaderKey.
- **Inputs**:
  - headerName (String): Parameter headerName
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for NormalizeHeaderKey.
- **Side Effects**: None

## Function: CampaignTableForSheet
- **Purpose**: Calculates or retrieves CampaignTableForSheet.
- **Inputs**:
  - ws (Worksheet): Parameter ws
- **Outputs**: ListObject
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CampaignTableForSheet.
- **Side Effects**: None

## Function: IsChecked
- **Purpose**: Calculates or retrieves IsChecked.
- **Inputs**:
  - inputValue (Variant): Parameter inputValue
- **Outputs**: Boolean
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for IsChecked.
- **Side Effects**: None

## Function: HasText
- **Purpose**: Calculates or retrieves HasText.
- **Inputs**:
  - inputValue (Variant): Parameter inputValue
- **Outputs**: Boolean
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for HasText.
- **Side Effects**: None

## Function: TextValue
- **Purpose**: Calculates or retrieves TextValue.
- **Inputs**:
  - inputValue (Variant): Parameter inputValue
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for TextValue.
- **Side Effects**: None

## Function: CheckedSymbol
- **Purpose**: Calculates or retrieves CheckedSymbol.
- **Inputs**:
  - None
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CheckedSymbol.
- **Side Effects**: None

## Function: UncheckedSymbol
- **Purpose**: Calculates or retrieves UncheckedSymbol.
- **Inputs**:
  - None
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for UncheckedSymbol.
- **Side Effects**: None

## Function: CreateBackupCopy
- **Purpose**: Calculates or retrieves CreateBackupCopy.
- **Inputs**:
  - wb (Workbook): Parameter wb
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CreateBackupCopy.
- **Side Effects**: None

## Function: CountBrokenReferences
- **Purpose**: Calculates or retrieves CountBrokenReferences.
- **Inputs**:
  - wb (Workbook): Parameter wb
- **Outputs**: Long
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CountBrokenReferences.
- **Side Effects**: None

## Function: ValidateWorkbookConfiguration
- **Purpose**: Calculates or retrieves ValidateWorkbookConfiguration.
- **Inputs**:
  - None
- **Outputs**: String
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ValidateWorkbookConfiguration.
- **Side Effects**: None

## Function: DashboardChartExists
- **Purpose**: Calculates or retrieves DashboardChartExists.
- **Inputs**:
  - chartName (String): Parameter chartName
- **Outputs**: Boolean
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for DashboardChartExists.
- **Side Effects**: None

## Function: ValidateCampaignTypeOptions
- **Purpose**: Entry point for ValidateCampaignTypeOptions.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ValidateCampaignTypeOptions.
- **Side Effects**: Modifies workbook state.

## Function: CampaignTypeDropdownIsConfigured
- **Purpose**: Calculates or retrieves CampaignTypeDropdownIsConfigured.
- **Inputs**:
  - rng (Range): Parameter rng
- **Outputs**: Boolean
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for CampaignTypeDropdownIsConfigured.
- **Side Effects**: None

## Function: RangeHasValidation
- **Purpose**: Calculates or retrieves RangeHasValidation.
- **Inputs**:
  - rng (Range): Parameter rng
- **Outputs**: Boolean
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for RangeHasValidation.
- **Side Effects**: None

## Function: RebuildMonthlyCalendars
- **Purpose**: Entry point for RebuildMonthlyCalendars.
- **Inputs**:
  - None
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for RebuildMonthlyCalendars.
- **Side Effects**: Modifies workbook state.

## Function: ClearCalendarSheet
- **Purpose**: Entry point for ClearCalendarSheet.
- **Inputs**:
  - ws (Worksheet): Parameter ws
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ClearCalendarSheet.
- **Side Effects**: Modifies workbook state.

## Function: AddDashboardCalendarLinks
- **Purpose**: Entry point for AddDashboardCalendarLinks.
- **Inputs**:
  - wb (Workbook): Parameter wb
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for AddDashboardCalendarLinks.
- **Side Effects**: Modifies workbook state.

## Function: GetOrCreateNotesInstructionSheet
- **Purpose**: Calculates or retrieves GetOrCreateNotesInstructionSheet.
- **Inputs**:
  - wb (Workbook): Parameter wb
- **Outputs**: Worksheet
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for GetOrCreateNotesInstructionSheet.
- **Side Effects**: None

## Function: ApplyWorkbookWrapText
- **Purpose**: Entry point for ApplyWorkbookWrapText.
- **Inputs**:
  - wb (Workbook): Parameter wb
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ApplyWorkbookWrapText.
- **Side Effects**: Modifies workbook state.

## Function: InstructionSheetIsReady
- **Purpose**: Calculates or retrieves InstructionSheetIsReady.
- **Inputs**:
  - None
- **Outputs**: Boolean
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for InstructionSheetIsReady.
- **Side Effects**: None

## Function: StyleCoreWorkbookSheets
- **Purpose**: Entry point for StyleCoreWorkbookSheets.
- **Inputs**:
  - wb (Workbook): Parameter wb
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for StyleCoreWorkbookSheets.
- **Side Effects**: Modifies workbook state.

## Function: OrderWorkbookSheets
- **Purpose**: Entry point for OrderWorkbookSheets.
- **Inputs**:
  - wb (Workbook): Parameter wb
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for OrderWorkbookSheets.
- **Side Effects**: Modifies workbook state.

## Function: UnfreezeWorkbookViews
- **Purpose**: Entry point for UnfreezeWorkbookViews.
- **Inputs**:
  - wb (Workbook): Parameter wb
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for UnfreezeWorkbookViews.
- **Side Effects**: Modifies workbook state.

## Function: ConfigureWorkbookViews
- **Purpose**: Entry point for ConfigureWorkbookViews.
- **Inputs**:
  - wb (Workbook): Parameter wb
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for ConfigureWorkbookViews.
- **Side Effects**: Modifies workbook state.

## Function: BuildNotesInstructionSheet
- **Purpose**: Entry point for BuildNotesInstructionSheet.
- **Inputs**:
  - wb (Workbook): Parameter wb
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for BuildNotesInstructionSheet.
- **Side Effects**: Modifies workbook state.

## Function: AddInstructionRow
- **Purpose**: Entry point for AddInstructionRow.
- **Inputs**:
  - ws (Worksheet): Parameter ws
  - rowNum (Long): Parameter rowNum
  - comp (String): Parameter comp
  - act (String): Parameter act
  - desc (String): Parameter desc
  - dep (String): Parameter dep
  - lim (String): Parameter lim
- **Outputs**: None
- **Dependencies**: Internal VBA components.
- **Behavior**: Executes logic for AddInstructionRow.
- **Side Effects**: Modifies workbook state.

# Module / File: Sheet3

# Module / File: Sheet6

# Module / File: Sheet7

# Module / File: Sheet8
