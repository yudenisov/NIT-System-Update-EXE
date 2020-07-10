Option Strict On
Imports System.IO
Imports System.Security
Imports System.Net
Imports Microsoft.VisualBasic.Devices
Imports NIT.SYSUPDATE.Library

Module Module1

    Sub Main()
        Install_WinMF51()
    End Sub

    ' *****************************************************************************
    '
    ' SUBROUTINE INSTALL_WinMF51
    ' This Subroutine Install Microsoft Windows Managements Framework 5.1
    ' on Local Computer
    '
    ' Parameters Read from Main INI File From Project Directory
    '
    ' PARAMETERS:	None
    '
    ' RETURNS:		None
    '
    ' *****************************************************************************
    Sub Install_WinMF51()
        Dim iResult As Integer
        iResult = 0
        Const Section_Name1 As String = "windows6.1-kb4019990-x86"
        Const Section_Name2 As String = "Win7-KB3191566-x86"
        Dim IWU As New InstallUpdateWindows.InstallUpdateWindows
        IWU.InstallWindowsUpdate(Section_Name1, UPDATE_WMF_SECTION)
        iResult = IWU.Is_IResult_0(iResult)
        IWU.InstallWindowsUpdate(Section_Name2, UPDATE_WMF_SECTION)
        iResult = IWU.Is_IResult_0(iResult)
        If iResult > 0 Then
            MsgBox("Error Installation of Windows Managements Framework 5.1. Program Terminated. Errno = " & Str(iResult))
        End If
    End Sub

End Module
