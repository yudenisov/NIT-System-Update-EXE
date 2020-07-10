Option Strict On
Imports System.IO
Imports System.Security
Imports System.Net
Imports Microsoft.VisualBasic.Devices
Imports NIT.SYSUPDATE.Library

Module Module1

    Sub Main()
        Install_Others()
    End Sub

    ' *****************************************************************************
    '
    ' SUBROUTINE INSTALL_Others
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
    Sub Install_Others()
        Dim iResult As Integer
        iResult = 0
        Const Section_Name1 As String = "chock.install"
        Const Section_Name2 As String = "ReverseMonitoringSetup"
        Const Section_Name3 As String = "stamp.check"
        Dim IWU As New InstallUpdateWindows.InstallUpdateWindows
        IWU.InstallBatFile(Section_Name1, UPDATE_OTHER_SECTION)
        iResult = IWU.Is_IResult_0(iResult)
        IWU.Install_ExeInnoSetup(Section_Name2, UPDATE_OTHER_SECTION)
        iResult = IWU.Is_IResult_0(iResult)
        IWU.InstallBatFile(Section_Name3, UPDATE_OTHER_SECTION)
        iResult = IWU.Is_IResult_0(iResult)
        If iResult > 0 Then
            MsgBox("Error Installation of Others Modules. Program Terminated. Errno = " & Str(iResult))
        Else
            MsgBox("Windows Updates With Success!")
        End If
    End Sub

End Module
