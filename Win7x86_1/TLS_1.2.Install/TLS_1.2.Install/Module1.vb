Option Strict On
Imports System.IO
Imports System.Security
Imports System.Net
Imports Microsoft.VisualBasic.Devices
Imports NIT.SYSUPDATE.Library

Module Module1

    Sub Main()
        Install_TLS12Updates()
    End Sub

    ' *****************************************************************************
    '
    ' SUBROUTINE INSTALL_TLS12UPDATES
    ' This Subroutine Install Microsoft TLS 1.2 Updates
    ' on Local Computer
    '
    ' Parameters Read from Main INI File From Project Directory
    '
    ' PARAMETERS:	None
    '
    ' RETURNS:		None
    '
    ' *****************************************************************************
    Sub Install_TLS12Updates()
        Dim iResult As Integer
        iResult = 0
        Const Section_Name1 As String = "windows6.1-kb3140245-x86"
        Dim IWU As New InstallUpdateWindows.InstallUpdateWindows
        IWU.InstallWindowsUpdate(Section_Name1, UPDATE_TLS_SECTION)
        iResult = IWU.Is_IResult_0(iResult)
        If iResult > 0 Then
            MsgBox("Error Installation of Windows Managements Framework 5.1. Program Terminated. Errno = " & Str(iResult))
        End If
    End Sub

End Module
