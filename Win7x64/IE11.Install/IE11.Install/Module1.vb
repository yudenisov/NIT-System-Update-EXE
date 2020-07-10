Option Strict On
Imports System.IO
Imports System.Security
Imports System.Net
Imports Microsoft.VisualBasic.Devices
Imports NIT.SYSUPDATE.Library

Module Module1

    Sub Main()
        Install_IE11Install()
    End Sub

    ' *****************************************************************************
    '
    ' SUBROUTINE INSTALL_IE11INSTALL
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
    Sub Install_IE11Install()
        Dim iResult As Integer
        iResult = 0
        Const Section_Name1 As String = "IE11-Windows6.1-x64-ru-ru"
        Dim IWU As New InstallUpdateWindows.InstallUpdateWindows
        IWU.Install_ExeInstShell(Section_Name1, UPDATE_IE11_SECTION2)
        iResult = IWU.Is_IResult_0(iResult)
        If iResult > 0 Then
            MsgBox("Error Installation of Windows Managements Framework 5.1. Program Terminated. Errno = " & Str(iResult))
        End If
    End Sub

End Module
