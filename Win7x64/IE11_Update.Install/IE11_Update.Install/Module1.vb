Option Strict On
Imports System.IO
Imports System.Security
Imports System.Net
Imports Microsoft.VisualBasic.Devices
Imports NIT.SYSUPDATE.Library

Module Module1

    Sub Main()
        Install_IE11Updates()
    End Sub

    ' *****************************************************************************
    '
    ' SUBROUTINE INSTALL_IE11UPDATES
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
    Sub Install_IE11Updates()
        Dim iResult As Integer
        iResult = 0
        Const Section_Name1 As String = "Windows6.1-KB2888049-x64"
        Const Section_Name2 As String = "Windows6.1-KB2882822-x64"
        Const Section_Name3 As String = "Windows6.1-KB2834140-v2-x64"
        Const Section_Name4 As String = "Windows6.1-KB2786081-x64"
        Const Section_Name5 As String = "Windows6.1-KB2731771-x64"
        Const Section_Name6 As String = "Windows6.1-KB2729094-v2-x64"
        Const Section_Name7 As String = "Windows6.1-KB2670838-x64"
        Const Section_Name8 As String = "Windows6.1-KB2533623-x64"
        Const Section_Name9 As String = "Windows6.1-KB2639308-x64"
        Dim IWU As New InstallUpdateWindows.InstallUpdateWindows
        IWU.InstallWindowsUpdate(Section_Name1, UPDATE_IE11_SECTION1)
        iResult = IWU.Is_IResult_0(iResult)
        IWU.InstallWindowsUpdate(Section_Name2, UPDATE_IE11_SECTION1)
        iResult = IWU.Is_IResult_0(iResult)
        IWU.InstallWindowsUpdate(Section_Name3, UPDATE_IE11_SECTION1)
        iResult = IWU.Is_IResult_0(iResult)
        IWU.InstallWindowsUpdate(Section_Name4, UPDATE_IE11_SECTION1)
        iResult = IWU.Is_IResult_0(iResult)
        IWU.InstallWindowsUpdate(Section_Name5, UPDATE_IE11_SECTION1)
        iResult = IWU.Is_IResult_0(iResult)
        IWU.InstallWindowsUpdate(Section_Name6, UPDATE_IE11_SECTION1)
        iResult = IWU.Is_IResult_0(iResult)
        IWU.InstallWindowsUpdate(Section_Name7, UPDATE_IE11_SECTION1)
        iResult = IWU.Is_IResult_0(iResult)
        IWU.InstallWindowsUpdate(Section_Name8, UPDATE_IE11_SECTION1)
        iResult = IWU.Is_IResult_0(iResult)
        IWU.InstallWindowsUpdate(Section_Name9, UPDATE_IE11_SECTION1)
        iResult = IWU.Is_IResult_0(iResult)
        If iResult > 0 Then
            MsgBox("Error Installation of Windows Managements Framework 5.1. Program Terminated. Errno = " & Str(iResult))
        End If
    End Sub

End Module
