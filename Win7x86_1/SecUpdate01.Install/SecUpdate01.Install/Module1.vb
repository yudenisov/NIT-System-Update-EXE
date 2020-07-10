Option Strict On
Imports System.IO
Imports System.Security
Imports System.Net
Imports Microsoft.VisualBasic.Devices
Imports NIT.SYSUPDATE.Library

Module Module1

    Sub Main()
        Install_SecUpdates01()
    End Sub

    ' *****************************************************************************
    '
    ' SUBROUTINE INSTALL_SecUpdates01
    ' This Subroutine Install Microsoft Internet Explorer 11 Updates
    ' on Local Computer
    '
    ' Parameters Read from Main INI File From Project Directory
    '
    ' PARAMETERS:	None
    '
    ' RETURNS:		None
    '
    ' *****************************************************************************
    Sub Install_SecUpdates01()
        Dim iResult As Integer
        iResult = 0
        If iResult > 0 Then
            MsgBox("Error Installation of Windows Secure Updates. Program Terminated. Errno = " & Str(iResult))
        End If
    End Sub

End Module
