Option Strict On
Imports NIT.SYSUPDATE.LIBRARY

Module Module1

    Sub Main(ByVal cmdArgs() As String)
        Dim IWU As New InstallUpdateWindows.InstallUpdateWindows
        Dim argNums As Integer
        argNums = cmdArgs.Length
        If argNums = 0 Then
            Console.WriteLine("No Arguments")
            IWU.Win10_WriteRegistry("")
        Else
            Console.WriteLine("Arguments {0}, First Argument: {1}", argNums, cmdArgs(0))
            IWU.Win10_WriteRegistry(cmdArgs(0))
        End If
        Shell("shutdown /r /t 00", vbHide)
    End Sub

End Module
