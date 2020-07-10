Option Strict On
Imports System.IO
Imports System.Security
Imports System.Net
Imports Microsoft.VisualBasic.Devices
Imports NIT.SYSUPDATE.Library



' *****************************************************************************
'
' MODULE SUBROUTINE_SYSUPDATE
'
' Main Module With Functions of NIT System Update Projects
'
' DEPENDS:	Constants.vb, Constants.sysupdate.vb, Subroutine_Module.vb
'		IniFile.vb
'
' *****************************************************************************
Module Subroutene_SysUpdate

    ' *****************************************************************************
    '
    ' DownloadFilesFromIntHttp( strFile, strURL, strPath )
    ' This Function Upload the File strFile from URL on HTTP/HTTPS Protocols
    ' and Save it on Local Computer to Path strPath
    ' Function Uses Objects "Microsoft.XMLHTTP" and "Adodb.Stream"
    '
    ' PARAMETERS:   strFile -- a File to be Downloaded (only name and extension)
    '               strURL -- an URL of the web-site, from which the File
    '               is Downloaded
    '               strPath -- a Place in a Windows Computer (Full path without slash)
    '               in which the File is Downloaded
    '
    ' RETURNS:      0 -- If File is Normally Downloaded and Created
    '               1 -- if File in Path strPath Can't Create
    '               2 -- If HTTP Response Not 200
    '
    ' *****************************************************************************

    Function DownloadFilesFromIntHttp(strFile As String, strURL As String, strPath As String) As Integer
        Dim strfileURL As String          'Full URL for file'
        Dim strLocal_Path As String       'full Path to local File'
        Dim intUploadFilesFromInt As Integer
        Dim Err01 As ArgumentException
        Dim Err02 As TimeoutException
        Dim Err03 As SecurityException
        Dim Err04 As WebException
        strfileURL = strURL & strFile
        strLocal_Path = strPath & "\" & strFile

        '**** Check if path is Exist ****'
        If Directory.Exists(strPath) Then
            intUploadFilesFromInt = 0
        Else
            intUploadFilesFromInt = 1
        End If

        ' If File Exsit Delete it
        If IO.File.Exists(strLocal_Path) Then
            IO.File.Delete(strLocal_Path)
            If File.Exists(strLocal_Path) Then
                intUploadFilesFromInt = 1
            Else
                intUploadFilesFromInt = 0
            End If
        End If

        ' **** Download File ****
        If intUploadFilesFromInt = 0 Then
            Try
                My.Computer.Network.DownloadFile(strfileURL, strLocal_Path)
                intUploadFilesFromInt = 0
                Exit Try
            Catch Err01
                MsgBox("Wrong Destination Path")
                intUploadFilesFromInt = 1
                Exit Try
            Catch Err02
                MsgBox("Time out Exception for URI: " & strfileURL)
                intUploadFilesFromInt = 2
                Exit Try
            Catch Err03
                MsgBox("Security Exception. Anonymous Authentication Denied")
                intUploadFilesFromInt = 2
                Exit Try
            Catch Err04
                MsgBox("Web Server Exception. Connection Refused")
                intUploadFilesFromInt = 2
                Exit Try
            End Try
        End If

        ' **** /Download File ****
        ' **** Check if Files is Downloaded **** '
        If Not File.Exists(strLocal_Path) And intUploadFilesFromInt = 0 Then
            intUploadFilesFromInt = 1
        End If
        ' **** /Check Path if Exist **** '
        DownloadFilesFromIntHttp = intUploadFilesFromInt
    End Function

    ' *****************************************************************************
    '
    ' FUNCTION CREATEDISTRIBFOLDER
    '
    ' This Function Creates 'Distrib' Folder at Project Root Directory or
    ' Deletes it if it is Exist
    '
    ' PARAMETERS:	szRoot -- the Project Root Directory
    '				blnFlag -- if vbFalse the Directory only Deleted
    '
    ' RETURNS:		0 -- the Directory is Created
    '				1 -- the Directory is Deleted, Not Created
    '				2 -- General Directory I/O Error
    '
    ' DEPENDS:		Constants.vb, Constants.sysupdate.vb
    '
    ' ****************************************************************************
    Function CreateDistribFolder(szRoot As String, blnFlag As Boolean) As Integer
        Const RMDIR As String = "rmdir /S /Q "
        Const MKDIR As String = "mkdir "
        Dim szPath As String
        Dim iResult As Integer
        If Not IO.Directory.Exists(szRoot) Then
            iResult = 2
        Else
            szPath = szRoot & "\" & DISTRIB_ROOT_FOLDER
            If IO.Directory.Exists(szPath) Then
                Shell(RMDIR & Chr(34) & szPath & Chr(34), vbHide)
                iResult = 1
            End If
            If blnFlag Then
                Shell(MKDIR & Chr(34) & szPath & Chr(34), vbHide)
                If Directory.Exists(szPath) Then
                    iResult = 0
                End If
            End If
        End If
        CreateDistribFolder = iResult
    End Function

    'Private Declare Sub Install_EXE Lib "NIT.SYSUPDATE.Library.dll" Alias "Install_ExeInnoSetup" (szSection As String, szSectionInstall As String)
    'Private Declare Function Is_Result Lib "NIT.SYSUPDATE.Library.dll" Alias "Is_Result" () As Integer
    'Private Declare Sub New_IWU Lib "NIT.SYSUPDATE.Library.dll" Alias "InstallUpdateWindows.New" ()
    'Private Declare Sub Set_Root Lib "NIT.SYSUPDATE.Library.dll" Alias "Set_Root" (ByVal strProjectRoot As String)
    ' *****************************************************************************
    '
    ' SUBROUTINE INSTALL_NETFRAMEWORK
    ' This Subroutine Install Microsoft .Net Framework on Local Computer
    '
    ' Parameters Read from Main INI File From Project Directory
    '
    ' PARAMETERS:	None
    '
    ' RETURNS:		None
    '
    ' *****************************************************************************

    Sub Install_NetFramework()
        Dim iResult As Integer
        Const Section_Name As String = "NET Version"
        Dim IWU As New InstallUpdateWindows.InstallUpdateWindows
        IWU.Install_ExeInstShell(Section_Name, UPDATE_NET_SECTION)
        iResult = IWU.Is_Result()
        If iResult > 0 Then
            MsgBox("Error Installation of .Net Framework 4.5. Program Terminated. Errno = " & Str(iResult))
        End If
    End Sub

End Module
