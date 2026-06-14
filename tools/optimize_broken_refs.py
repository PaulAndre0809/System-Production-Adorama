import win32com.client as win32
from pathlib import Path
import time
import pythoncom
import pywintypes

ORIGINAL_FUNCTION = """Private Function CountBrokenReferences(ByVal wb As Workbook) As Long

    Dim ws As Worksheet

    Dim formulas As Range

    Dim cell As Range

    Dim nm As Name

    Dim total As Long

    For Each ws In wb.Worksheets

        Set formulas = Nothing

        On Error Resume Next

        Set formulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)

        On Error GoTo 0

        If Not formulas Is Nothing Then

            For Each cell In formulas.Cells

                If InStr(1, CStr(cell.Formula), "#REF!", vbTextCompare) > 0 Then

                    total = total + 1

                End If

            Next cell

        End If

    Next ws

    For Each nm In wb.Names

        On Error Resume Next

        If InStr(1, CStr(nm.RefersTo), "#REF!", vbTextCompare) > 0 Then

            total = total + 1

        End If

        On Error GoTo 0

    Next nm

    CountBrokenReferences = total

End Function"""

OPTIMIZED_FUNCTION = """Private Function CountBrokenReferences(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim formulaArray As Variant
    Dim r As Long, c As Long
    Dim nm As Name
    Dim total As Long
    For Each ws In wb.Worksheets
        On Error Resume Next
        formulaArray = ws.UsedRange.Formula
        On Error GoTo 0
        If Not IsEmpty(formulaArray) Then
            If IsArray(formulaArray) Then
                For r = 1 To UBound(formulaArray, 1)
                    For c = 1 To UBound(formulaArray, 2)
                        If InStr(1, CStr(formulaArray(r, c)), "#REF!", vbTextCompare) > 0 Then
                            total = total + 1
                        End If
                    Next c
                Next r
            Else
                If InStr(1, CStr(formulaArray), "#REF!", vbTextCompare) > 0 Then
                    total = total + 1
                End If
            End If
        End If
    Next ws
    For Each nm In wb.Names
        On Error Resume Next
        If InStr(1, CStr(nm.RefersTo), "#REF!", vbTextCompare) > 0 Then
            total = total + 1
        End If
        On Error GoTo 0
    Next nm
    CountBrokenReferences = total
End Function"""

def retry(func, attempts=10, delay=0.5):
    for i in range(attempts):
        try:
            pythoncom.PumpWaitingMessages()
            return func()
        except pywintypes.com_error as e:
            if i == attempts - 1:
                raise
            time.sleep(delay)

def main():
    repo_dir = Path(__file__).resolve().parents[1]
    wb_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker.xlsm"
    template_path = repo_dir / "Production Tracker" / "Email & SMS Campaign Tracker Template.xlsm"
    
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1
        
        for p in [wb_path, template_path]:
            print(f"Opening workbook: {p.name}")
            wb = retry(lambda: excel.Workbooks.Open(str(p), UpdateLinks=0))
            
            try:
                comp = wb.VBProject.VBComponents("modEmailProductionTracker")
                count = comp.CodeModule.CountOfLines
                if count > 0:
                    code = comp.CodeModule.Lines(1, count)
                    
                    # Normalize double lines in ORIGINAL_FUNCTION query for matching
                    # We can just look for parts of CountBrokenReferences to replace
                    # Since it might have double lines or single lines
                    # Let's find the start of the function and the end
                    start_str = "Private Function CountBrokenReferences("
                    end_str = "End Function"
                    
                    start_idx = code.find(start_str)
                    if start_idx != -1:
                        # Find the first End Function after start_idx
                        end_idx = code.find(end_str, start_idx)
                        if end_idx != -1:
                            end_idx += len(end_str)
                            old_func = code[start_idx:end_idx]
                            
                            code = code[:start_idx] + OPTIMIZED_FUNCTION + code[end_idx:]
                            
                            comp.CodeModule.DeleteLines(1, count)
                            comp.CodeModule.AddFromString(code)
                            wb.Save()
                            print(f"Successfully optimized CountBrokenReferences in {p.name}")
                        else:
                            print(f"End Function not found after CountBrokenReferences in {p.name}")
                    else:
                        print(f"CountBrokenReferences not found in {p.name}")
            except Exception as e:
                print(f"Failed to process {p.name}: {e}")
                
            wb.Close(SaveChanges=False)
            
    except Exception as e:
        print(f"Error: {e}")
    finally:
        if excel:
            try:
                excel.Quit()
            except:
                pass

if __name__ == "__main__":
    main()
