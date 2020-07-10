Module Constants
    Public Const FIRMNAME As String = "New_Internet_Technologies"
    Public Const PROJECTNAME As String = "Win7x86"
    Public Const INI_FILE1 As String = "test.ini"
    Public Const INI_MAIN As String = "main.ini"
    Public Const ROOT_DIRNAME As String = "C:\ProgramData"

    Public Function IsINIFile_Exist(szINI_File) As Integer
        Dim fso
        Dim szPath As String
        szPath = ROOT_DIRNAME
        szPath = szPath & "\" & FIRMNAME & "\" & PROJECTNAME & "\" & szINI_File
        fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(szPath) Then
            IsINIFile_Exist = 0
        Else
            IsINIFile_Exist = 1
        End If
    End Function

    Public Function Create_ProjectAppData_Folders() As Integer
        Dim fso, fdir
        Dim szAppDataPath
        szAppDataPath = ROOT_DIRNAME
        fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(szAppDataPath) Then
            MsgBox("Error: System Path " & szAppDataPath & " not Found!", MsgBoxStyle.OkOnly And MsgBoxStyle.Critical)
            Create_ProjectAppData_Folders = 2
        Else
            szAppDataPath = szAppDataPath & "\" & FIRMNAME
            If Not fso.FolderExists(szAppDataPath) Then
                fdir = fso.CreateFolder(szAppDataPath)
            End If
            szAppDataPath = szAppDataPath & "\" & PROJECTNAME
            If Not fso.FolderExists(szAppDataPath) Then
                fdir = fso.CreateFolder(szAppDataPath)
                If fso.FolderExists(szAppDataPath) Then
                    Create_ProjectAppData_Folders = 0
                Else
                    Create_ProjectAppData_Folders = 2
                End If
            Else
                Create_ProjectAppData_Folders = 1
            End If
        End If
    End Function
End Module
