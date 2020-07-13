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
Dim REGISTRY_KEY_EX
REGISTRY_KEY_EX = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
Set objShell = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
CurrentPath = fso.GetParentFolderName(WScript.ScriptFullName)
STRFULLREGPROGRAM = CurrentPath & "\regprogram.exe"
If fso.FileExists(STRFULLREGPROGRAM) Then
  WshShell.RegWrite REGISTRY_KEY_EX & "\" & "1", Chr(34) & STRFULLREGPROGRAM & Chr(34), "REG_SZ" 
Else
  WScript.Echo "Installation Abnormal! Quit." & vbCrLf & "Инсталляция неуспешная! Выход."  & vbCrLf
End If


