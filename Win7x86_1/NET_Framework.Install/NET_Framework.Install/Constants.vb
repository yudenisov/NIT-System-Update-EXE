Option Strict On
Imports System.IO

Module Constants
    Public Const FIRMNAME As String = "NIT"
    Public Const PROJECTNAME As String = "7x86"
    Public Const INI_FILE1 As String = "test.ini"
    Public Const INI_MAIN As String = "main.ini"
    Public Const ROOT_DIRNAME As String = "C:\ProgramData"

    Public Function IsINIFile_Exist(ByVal szINI_File As String) As Integer
        Dim szPath As String
        szPath = ROOT_DIRNAME
        szPath = szPath & "\" & FIRMNAME & "\" & PROJECTNAME & "\" & szINI_File
        If IO.File.Exists(szPath) Then
            IsINIFile_Exist = 0
        Else
            IsINIFile_Exist = 1
        End If
    End Function

    Public Function Create_ProjectAppData_Folders() As Integer
        Dim szAppDataPath As String
        Dim fdir As DirectoryInfo
        szAppDataPath = ROOT_DIRNAME
        If IO.Directory.Exists(szAppDataPath) Then
            szAppDataPath = szAppDataPath & "\" & FIRMNAME
            If Not Directory.Exists(szAppDataPath) Then
                fdir = Directory.CreateDirectory(szAppDataPath)
            End If
            szAppDataPath = szAppDataPath & "\" & PROJECTNAME
            If Not Directory.Exists(szAppDataPath) Then
                fdir = Directory.CreateDirectory(szAppDataPath)
                If Directory.Exists(szAppDataPath) Then
                    Create_ProjectAppData_Folders = 0
                Else
                    Create_ProjectAppData_Folders = 2
                End If
            Else
                Create_ProjectAppData_Folders = 1
            End If
        Else
            MsgBox("Error: System Path " & szAppDataPath & " not Found!", MsgBoxStyle.OkOnly And MsgBoxStyle.Critical)
            Create_ProjectAppData_Folders = 2
        End If
    End Function
End Module
