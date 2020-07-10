Option Explicit
' *****************************************************************************
'
' Script runprogram.vbs
' This Script Runs the regprogram.exe Program with Elevated Privileges
'
' PARAMETERS:	1 (REQUARED) Full Path Neme to regprogram.exe
'		2 (OPTIONAL) the Section of main.ini File
'
' RETURNS:	NONE
'
' *****************************************************************************
Dim objArgs, objShell
Set objShell = CreateObject("Shell.Application")
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
  WScript.Echo "No Arguments"
End If
If objArgs.Count = 1 Then
  objShell.ShellExecute objArgs(0), "", "", "runas", 1
End If
If objArgs.Count = 2 Then
  objShell.ShellExecute objArgs(0), objArgs(1), "", "runas", 1
End If
If objArgs.Count > 2 Then
  WScript.Echo "Extra Arguments"
End If

