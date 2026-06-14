import sys
import time
import win32com.client as win32
from pathlib import Path
import shutil
import tempfile
import os

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    debug_log_path = repo_dir / "tools" / "debug_log.txt"
    
    # Clean up previous log if any
    if debug_log_path.exists():
        try:
            debug_log_path.unlink()
        except:
            pass
            
    # Copy workbook to temp
    temp_dir = Path(tempfile.mkdtemp(prefix="debug_hang_"))
    temp_wb_path = temp_dir / wb_path.name
    shutil.copy2(wb_path, temp_wb_path)
    print(f"Copied workbook to {temp_wb_path}")
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1
        
        print("Opening copy to insert debugging subroutine...")
        wb = excel.Workbooks.Open(str(temp_wb_path), UpdateLinks=0)
        
        comp = wb.VBProject.VBComponents("modEmailProductionTracker")
        
        log_file_vba_escaped = str(debug_log_path).replace("\\", "\\\\")
        debug_sub = f"""
Private Sub LogDebugStep(ByVal msg As String)
    Dim f As Long
    f = FreeFile
    Open "{log_file_vba_escaped}" For Append As #f
    Print #f, msg
    Close #f
End Sub

Public Sub DebugApplyStepByStep()
    Dim lo As ListObject, smsTable As ListObject
    
    LogDebugStep "START"
    
    LogDebugStep "1. EnsureCampaignSheets starting..."
    EnsureCampaignSheets
    LogDebugStep "1. EnsureCampaignSheets finished."
    
    LogDebugStep "2. Getting tables..."
    Set lo = GetInventoryTable()
    Set smsTable = GetSmsTable()
    LogDebugStep "2. Getting tables finished."
    
    LogDebugStep "3. Rename Bluecore columns..."
    EnsureRenamedColumn lo, "Bluecore Link", "Bluecore/Attentive Link"
    EnsureRenamedColumn smsTable, "Bluecore Link", "Bluecore/Attentive Link"
    LogDebugStep "3. Rename Bluecore columns finished."
    
    LogDebugStep "4. EnsureNotesColumn..."
    EnsureNotesColumn lo
    EnsureNotesColumn smsTable
    LogDebugStep "4. EnsureNotesColumn finished."
    
    LogDebugStep "5. ConfigureChecklistColumn UTM..."
    ConfigureChecklistColumn lo, "Campaign Name and UTM Parameter (Source Code)"
    LogDebugStep "5. ConfigureChecklistColumn CreativeBrief..."
    ConfigureChecklistColumn lo, "Creative Brief, SL & PH"
    LogDebugStep "5. ConfigureChecklistColumn SKUs..."
    ConfigureChecklistColumn lo, "SKUs"
    LogDebugStep "5. ConfigureChecklistColumn In-Design..."
    ConfigureChecklistColumn lo, "In-Design"
    LogDebugStep "5. ConfigureChecklistColumn Build..."
    ConfigureChecklistColumn lo, "Build, QA"
    LogDebugStep "5. ConfigureChecklistColumn Route..."
    ConfigureChecklistColumn lo, "Route"
    LogDebugStep "5. ConfigureChecklistColumn Approval..."
    ConfigureChecklistColumn lo, "Approval"
    LogDebugStep "5. ConfigureChecklistColumn Segments..."
    ConfigureChecklistColumn lo, "Segments"
    LogDebugStep "5. ConfigureChecklistColumn Email finished."
    
    LogDebugStep "6. ConfigureChecklistColumn SMS Options..."
    ConfigureChecklistColumn smsTable, "Send SMS Options"
    LogDebugStep "6. ConfigureChecklistColumn SMS Test..."
    ConfigureChecklistColumn smsTable, "Send Test"
    LogDebugStep "6. ConfigureChecklistColumn SMS Approval..."
    ConfigureChecklistColumn smsTable, "Approval"
    LogDebugStep "6. ConfigureChecklistColumn SMS Segments..."
    ConfigureChecklistColumn smsTable, "Segments"
    LogDebugStep "6. ConfigureChecklistColumn SMS finished."
    
    LogDebugStep "7. ConfigureOwnerColumn Email..."
    ConfigureOwnerColumn lo
    LogDebugStep "8. EnsureUpdatedByColumn Email..."
    EnsureUpdatedByColumn lo
    LogDebugStep "9. ConfigureCampaignTypeColumn Email..."
    ConfigureCampaignTypeColumn lo
    LogDebugStep "10. FormatSendDateColumn Email..."
    FormatSendDateColumn lo
    LogDebugStep "11. FormatSendTimeColumn Email..."
    FormatSendTimeColumn lo
    LogDebugStep "12. ApplyCalculatedColumns Email..."
    ApplyCalculatedColumns lo
    LogDebugStep "12. ApplyCalculatedColumns Email finished."
    
    LogDebugStep "13. ConfigureOwnerColumn SMS..."
    ConfigureOwnerColumn smsTable
    LogDebugStep "14. EnsureUpdatedByColumn SMS..."
    EnsureUpdatedByColumn smsTable
    LogDebugStep "15. ConfigureCampaignTypeColumn SMS..."
    ConfigureCampaignTypeColumn smsTable
    LogDebugStep "16. FormatSendDateColumn SMS..."
    FormatSendDateColumn smsTable
    LogDebugStep "17. FormatSendTimeColumn SMS..."
    FormatSendTimeColumn smsTable
    LogDebugStep "18. ApplyCalculatedColumns SMS..."
    ApplyCalculatedColumns smsTable
    LogDebugStep "18. ApplyCalculatedColumns SMS finished."
    
    LogDebugStep "19. RebuildMonthlyCalendars..."
    RebuildMonthlyCalendars
    LogDebugStep "19. RebuildMonthlyCalendars finished."
    
    LogDebugStep "20. RefreshDashboard..."
    RefreshDashboard
    LogDebugStep "20. RefreshDashboard finished."
    
    LogDebugStep "21. UpdateCalendarTabs..."
    UpdateCalendarTabs
    LogDebugStep "21. UpdateCalendarTabs finished."
    
    LogDebugStep "22. StyleCoreWorkbookSheets..."
    StyleCoreWorkbookSheets ThisWorkbook
    LogDebugStep "22. StyleCoreWorkbookSheets finished."
    
    LogDebugStep "SUCCESS"
End Sub
"""
        comp.CodeModule.AddFromString(debug_sub)
        wb.Save()
        print("Debugging subroutine added and workbook saved.")
        
        print("Running DebugApplyStepByStep...")
        excel.Run(f"'{wb.Name}'!DebugApplyStepByStep")
        print("VBA execution finished normally.")
        
        if debug_log_path.exists():
            print("Log contents:")
            print(debug_log_path.read_text())
            
        wb.Close(SaveChanges=False)
    except Exception as e:
        print(f"Error occurred: {e}")
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == "__main__":
    main()
