"""Embed modDynamicHighlighting VBA module into the Campaign Tracker workbooks and update RefreshDashboard."""
import sys
import time
from pathlib import Path
import pythoncom
import pywintypes
import win32com.client as win32

def retry(label, func, attempts=20, delay=0.75):
    last_error = None
    for _ in range(attempts):
        try:
            pythoncom.PumpWaitingMessages()
            return func()
        except pywintypes.com_error as exc:
            last_error = exc
            if exc.args and exc.args[0] in (-2147418111, -2147417846):
                time.sleep(delay)
                continue
            raise
    raise RuntimeError(f"{label} failed repeatedly: {last_error}")

def embed_vba_in_workbook(excel, file_path: Path, vba_code: str):
    print(f"Opening workbook: {file_path}")
    workbook = retry(
        "open workbook",
        lambda: excel.Workbooks.Open(str(file_path), UpdateLinks=0, ReadOnly=False, AddToMru=False)
    )
    
    vb_project = workbook.VBProject
    
    # Check if modDynamicHighlighting already exists; if so, remove it to prevent duplicates
    mod_name = "modDynamicHighlighting"
    existing_comp = None
    for i in range(1, vb_project.VBComponents.Count + 1):
        comp = vb_project.VBComponents(i)
        if comp.Name == mod_name:
            existing_comp = comp
            break
            
    if existing_comp is not None:
        print(f"Removing existing module '{mod_name}'...")
        vb_project.VBComponents.Remove(existing_comp)
        
    # Add new module
    print(f"Adding new module '{mod_name}'...")
    new_comp = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
    new_comp.Name = mod_name
    new_comp.CodeModule.AddFromString(vba_code)
    
    # Now find modEmailProductionTracker to auto-integrate RefreshDashboard call
    tracker_comp = None
    for i in range(1, vb_project.VBComponents.Count + 1):
        comp = vb_project.VBComponents(i)
        if comp.Name == "modEmailProductionTracker":
            tracker_comp = comp
            break
            
    if tracker_comp is None:
        print("Warning: 'modEmailProductionTracker' module not found. Automation integration skipped.")
    else:
        print("Integrating auto-refresh into 'modEmailProductionTracker'...")
        count = tracker_comp.CodeModule.CountOfLines
        if count > 0:
            code = tracker_comp.CodeModule.Lines(1, count)
            
            # Locate StyleDashboard and ApplyDashboardNativeFormulas calls in RefreshDashboard
            target_pattern = "ApplyDashboardNativeFormulas wsD, dashboardTable"
            highlight_call = "    Call ApplyDynamicScheduleFilterAndHighlight(dashboardTable)"
            
            if target_pattern in code:
                if highlight_call not in code:
                    # Insert the call right after the target pattern
                    new_code = code.replace(
                        target_pattern,
                        f"{target_pattern}\n{highlight_call}"
                    )
                    tracker_comp.CodeModule.DeleteLines(1, count)
                    tracker_comp.CodeModule.AddFromString(new_code)
                    print("Auto-refresh call successfully integrated into RefreshDashboard!")
                else:
                    print("Auto-refresh call already present in VBA code.")
            else:
                print("Warning: Could not locate RefreshDashboard injection point in VBA code.")
                
    # Save and close workbook
    print(f"Saving and closing workbook: {file_path}")
    workbook.Save()
    workbook.Close(SaveChanges=True)

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    vba_file = repo_dir / "Production Tracker" / "dynamic_highlighting.vba"
    
    if not vba_file.exists():
        print(f"VBA script source not found at: {vba_file}")
        sys.exit(1)
        
    with open(vba_file, "r") as f:
        vba_code = f.read()
        
    workbooks = [
        repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm",
        repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker Template.xlsm"
    ]
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False
        try:
            excel.AutomationSecurity = 1 # xlSecurityLevelNone to permit macro modification
        except Exception:
            pass
            
        for wb_path in workbooks:
            if wb_path.exists():
                embed_vba_in_workbook(excel, wb_path, vba_code)
            else:
                print(f"Workbook not found: {wb_path}")
                
        print("\nAll workbooks successfully updated!")
    except Exception as e:
        print(f"\nAn error occurred: {e}", file=sys.stderr)
        sys.exit(1)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

if __name__ == "__main__":
    main()
