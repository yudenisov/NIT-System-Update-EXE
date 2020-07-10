Option Explicit
' *****************************************************************************
'
' Script regprogram.vbs
' This Script Runs the regprogram.exe Program with Elevated Privileges
'
' PARAMETERS:	NONE
'
' RETURNS:	NONE
'
' *****************************************************************************
Dim objShell, fso, WshShell
Dim STRFULLREGPROGRAM, CurrentPath
Set objShell = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
CurrentPath = fso.GetParentFolderName(WScript.ScriptFullName)
STRFULLREGPROGRAM = CurrentPath & "\regprogram.vbs"
If fso.FileExists(STRFULLREGPROGRAM) Then
  objShell.ShellExecute "C:\Windows\System32\cscript.exe", "//Nologo " & Chr(34) & STRFULLREGPROGRAM & Chr(34), "", "runas", 1
Else
  WScript.Echo "Installation Abnormal! Quit." & vbCrLf & "Инсталляция неуспешная! Выход."  & vbCrLf
End If


