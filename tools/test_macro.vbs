Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
objExcel.AutomationSecurity = 1

Set fso = CreateObject("Scripting.FileSystemObject")
wbPath = fso.GetAbsolutePathName("Production Tracker\Email & SMS Campaign Tracker.xlsm")

On Error Resume Next
Set objWorkbook = objExcel.Workbooks.Open(wbPath, False, False)
If Err.Number <> 0 Then
    WScript.Echo "Error opening workbook: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Echo "Running RefreshProductionStatus..."
On Error Resume Next
objExcel.Run "'" & objWorkbook.Name & "'!RefreshProductionStatus"
If Err.Number <> 0 Then
    WScript.Echo "Macro failed: " & CStr(Err.Number) & " - " & Err.Description
Else
    WScript.Echo "Macro finished successfully!"
End If
On Error GoTo 0

objWorkbook.Close False
objExcel.Quit
WScript.Echo "Done."
