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
        Dim fso, xmlHttp, adoStream As Object
        fso = CreateObject("Scripting.FileSystemObject")
        xmlHttp = CreateObject("Microsoft.XMLHTTP")
        adoStream = CreateObject("Adodb.Stream")
        Dim strfileURL As String          'Full URL for file'
        Dim strLocal_Path As String       'full Path to local File'
        Dim intUploadFilesFromInt As Integer
        Dim blnExistRemoteFile As Boolean
        strfileURL = strURL & strFile
        strLocal_Path = strPath & "\" & strFile

        '**** Check if path is Exist ****'
        If fso.FolderExists(strPath) Then
            intUploadFilesFromInt = 0
        Else
            intUploadFilesFromInt = 1
        End If

        ' If File Exsit Delete it
        If fso.FileExists(strLocal_Path) Then
            fso.DeleteFile(strLocal_Path, "Force")
            If fso.FileExists(strLocal_Path) Then
                intUploadFilesFromInt = 1
            Else
                intUploadFilesFromInt = 0
            End If
        End If

        ' **** Download File ****
        'MsgBox strfileURL
        xmlHttp.Open("GET", strfileURL, False)
        xmlHttp.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Safari/537.36")
        xmlHttp.Send
        'MsgBox xmlHttp.statusText
        If xmlHttp.Status = 200 And intUploadFilesFromInt = 0 Then
            blnExistRemoteFile = vbTrue
        Else
            blnExistRemoteFile = vbFalse
            intUploadFilesFromInt = 2
            xmlHttp.Abort
        End If
        If blnExistRemoteFile Then
            adoStream.Type = 1
            adoStream.Mode = 3
            adoStream.Open
            adoStream.Write(xmlHttp.responseBody)
            adoStream.SaveToFile(strLocal_Path, 2)
            '       **** /Download File ****

            adoStream.Close
            xmlHttp.Abort

            ' **** Check if Files is Downloaded **** '
            If Not fso.FileExists(strLocal_Path) And intUploadFilesFromInt = 0 Then
                intUploadFilesFromInt = 1
            End If
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
        Dim WshShell, fso As Object
        Dim szPath As String
        Dim iResult As Integer
        WshShell = CreateObject("WScript.Shell")
        fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(szRoot) Then
            iResult = 2
        Else
            szPath = szRoot & "\" & DISTRIB_ROOT_FOLDER
            If fso.FolderExists(szPath) Then
                WshShell.Run(RMDIR & Chr(34) & szPath & Chr(34), 0, True)
                iResult = 1
            End If
            If blnFlag Then
                WshShell.Run(MKDIR & Chr(34) & szPath & Chr(34), 0, True)
                If fso.FolderExists(szPath) Then
                    iResult = 0
                End If
            End If
        End If
        CreateDistribFolder = iResult
    End Function

    ' *****************************************************************************
    '
    ' SUBROUTINE INSTALL_NETFRAMEWORK
    ' This Subroutine Install Microsoft .Net Framework on Local Computer
    '
    ' Parameters Read from Main INI File From Projet Directory
    '
    ' PARAMETERS:	szSection_Name -- Name of Section at INI File
    '
    ' RETURNS:		None
    '
    ' *****************************************************************************

    Sub Install_NetFramework()
        Dim iResult As Integer
        Const Section_Name As String = "NET Version"
        iResult = InstallWindowsUpdate(Section_Name, UPDATE_NET_SECTION)
        If iResult = 1 Then
            MsgBox("Error Installation of .Net Framework 4.5. Program Terminated. Errno = " & Str(iResult))
        End If
    End Sub

End Module
