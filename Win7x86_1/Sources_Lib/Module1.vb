Module Module1

    Sub Main()
        Dim strProjectRoot As String
        Dim iResult As Integer
        strProjectRoot = ROOT_DIRNAME & "\" & FIRMNAME & "\" & PROJECTNAME
        iResult = CreateDistribFolder(strProjectRoot, True)
        If iResult > 0 Then
            MsgBox("Error Created Distrib Folder. Program Terminated. Errno = " & Str(iResult))
        Else
            'MsgBox("Distrib Folder is Created!")
            Install_NetFramework()
        End If
    End Sub

End Module
