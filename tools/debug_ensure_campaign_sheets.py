import win32com.client as win32
from pathlib import Path
import time

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = True
        excel.EnableEvents = True
        
        print("Opening workbook...")
        wb = excel.Workbooks.Open(str(wb_path), UpdateLinks=0)
        
        print("Checking worksheets...")
        email_sheet = None
        try:
            email_sheet = wb.Worksheets("Email Campaigns")
            print("  Found 'Email Campaigns' sheet")
        except Exception as e:
            print(f"  'Email Campaigns' sheet not found: {e}")
            
        if email_sheet is None:
            try:
                email_sheet = wb.Worksheets("Production Inventory")
                email_sheet.Name = "Email Campaigns"
                print("  Renamed 'Production Inventory' to 'Email Campaigns'")
            except Exception as e:
                print(f"  'Production Inventory' not found: {e}")
                
        if email_sheet is None:
            print("  ERROR: Email Campaigns sheet is missing!")
            return
            
        print("Checking Email table...")
        lo_email = None
        try:
            lo_email = email_sheet.ListObjects("EmailCampaignsTable")
            print("  Found 'EmailCampaignsTable'")
        except Exception as e:
            print(f"  'EmailCampaignsTable' not found: {e}")
            
        if lo_email is None:
            try:
                lo_email = email_sheet.ListObjects("ProductionInventoryTable")
                lo_email.Name = "EmailCampaignsTable"
                print("  Renamed table to 'EmailCampaignsTable'")
            except Exception as e:
                print(f"  'ProductionInventoryTable' not found: {e}")
                
        print("Checking SMS sheet...")
        sms_sheet = None
        try:
            sms_sheet = wb.Worksheets("SMS Campaigns")
            print("  Found 'SMS Campaigns' sheet")
        except Exception as e:
            print(f"  'SMS Campaigns' sheet not found: {e}")
            
        if sms_sheet is None:
            sms_sheet = wb.Worksheets.Add(After=email_sheet)
            sms_sheet.Name = "SMS Campaigns"
            print("  Created 'SMS Campaigns' sheet")
            
        print("Checking SMS table...")
        lo_sms = None
        try:
            lo_sms = sms_sheet.ListObjects("SMSCampaignsTable")
            print("  Found 'SMSCampaignsTable'")
        except Exception as e:
            print(f"  'SMSCampaignsTable' not found: {e}")
            
        if lo_sms is None:
            print("  SMS table not found, checking if sheet is empty...")
            used_range = sms_sheet.UsedRange
            # CountA
            count_a = excel.WorksheetFunction.CountA(used_range)
            print(f"  CountA of SMS used range: {count_a}")
            
            if count_a > 0:
                print("  ERROR: SMS Campaigns contains data but no SMSCampaignsTable.")
                return
                
            print("  Creating SMS Campaigns table headers...")
            headers = [
                "Send Date", "Send Time", "Campaign Name", "Campaign Type",
                "Current Stage", "Owner", "Send SMS Options", "Send Test",
                "Approval", "Segments", "Jira Link", "ClickUp Link",
                "Bluecore/Attentive Link", "Est. Audience", "Delivered",
                "Last Updated", "Last Updated By", "Notes"
            ]
            for i, h in enumerate(headers):
                sms_sheet.Cells(1, i + 1).Value = h
                
            print("  Adding SMS ListObject...")
            lo_sms = sms_sheet.ListObjects.Add(
                SourceType=1, # xlSrcRange
                Source=sms_sheet.Range("A1:R201"),
                XlListObjectHasHeaders=1 # xlYes
            )
            lo_sms.Name = "SMSCampaignsTable"
            print("  SMS Campaigns table created successfully")
            
        print("Setting SMS table style...")
        lo_sms.TableStyle = "TableStyleMedium2"
        print("EnsureCampaignSheets steps finished successfully in Python!")
        
        wb.Close(SaveChanges=False)
    except Exception as e:
        print(f"Error during execution: {e}")
    finally:
        if excel:
            excel.Quit()

if __name__ == "__main__":
    main()
