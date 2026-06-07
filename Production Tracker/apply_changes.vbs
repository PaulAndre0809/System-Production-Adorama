Option Explicit

Dim fso
Dim shell
Dim scriptDirectory
Dim repositoryDirectory
Dim pythonPath
Dim scriptPath
Dim command
Dim exitCode

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
repositoryDirectory = fso.GetParentFolderName(scriptDirectory)
pythonPath = repositoryDirectory & "\.venv\Scripts\python.exe"
scriptPath = scriptDirectory & "\apply_changes.py"

If Not fso.FileExists(pythonPath) Then
    WScript.Echo "ERROR: Python environment not found: " & pythonPath
    WScript.Quit 1
End If

command = """" & pythonPath & """ """ & scriptPath & """ --visible"
exitCode = shell.Run(command, 1, True)
WScript.Quit exitCode
