' *****************************************************************************
'
' SUBROUTINE Module
' Library with Difference Functions for NIT Projects
'
' *****************************************************************************

Module Subroutine_Module

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
    ' FUNCTION INSTALLBATFILE
    '
    ' This Function Install Windows Update File.
    ' Parameters Loads from INI File
    '
    ' PARAMETERS:	szSection -- the Update File Section
    '				szSectionInstall -- the INI File Section
    '
    ' RETURNS:		0 -- the File is Installed With Success
    '				1 -- Some Errors Occur
	'				2 -- Execution Type Mismatch
    '
    ' DEPENDS:		Constants.vb, Constants.sysupdate.vb
    '
    ' ****************************************************************************
    Function InstallBatFile( szSection As String, szSectionInstall As String) As Integer
        Dim WshShell, fso As Object
        Dim iResult As Integer
        Dim szRootDir As String
        Dim szMainINI As String
        Dim szINI As String
        szRootDir = ROOT_DIRNAME & "\" & FIRMNAME & "\" & PROJECTNAME
		Console.WriteLine("Program Root Directory = {0}", szRootDir)
        szMainINI = szRootDir & "\" & INI_MAINC
		Console.WriteLine( "Main INI File = {0}", szMainINI )
        fso = CreateObject( "Scripting.FileSystemObject" )
        If fso.FileExists(szMainINI) Then
            Dim MainIniFile As New Ini.IniFile(szMainINI)
            szINI = MainIniFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
            szINI = szRootDir & "\" & szINI
			Console.WriteLine("INI FIle = {0}", szINI)
            If fso.FileExists(szINI) Then
                Dim iniFile As New Ini.IniFile(szINI)
                Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
				Dim szBuild As String
                szPath = szRootDir & "\" & DISTRIB_ROOT_FOLDER
                szHost = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                szPref = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                szPort = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                szUser = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                szPass = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                szFileName = iniFile.ReadKey(szSection, "filename")
                szExtension = iniFile.ReadKey(szSection, "extension")
                If Not StrComp(szUser, "", 0) Or Not StrComp(szPref, "http", 0) Then
                    szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                    szLocalFile = szFileName & szExtension
					Console.WriteLine("URL Path = {0}", szURL )
					Console.WriteLine("Local Path = {0}", szPath )
					Console.WriteLine("File: {0}", szLocalFile)
                    If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                        iResult = 1
                    Else
						szBuild = iniFile.ReadKey( szSection, BUILD_PROGRAM )
						If StrComp(szExtension, ".bat", 0) = 0 Or StrComp(szExtension, ".cmd", 0) = 0 And StrComp(szBuild, BUILD_BAT, 0) = 0 Then
							Console.WriteLine("Build Program = {0}", BUILD_BAT )
							WshShell = CreateObject("WScript.Shell")
							MsgBox("cmd /c cd /d " & Chr(34) & szPath & Chr(34) & " && " & Chr(34) & szPath & "\" & szLocalFile & Chr(34))
							'WshShell.Run("cmd /c cd /d " & Chr(34) & szPath & Chr(34) & " && " & Chr(34) & szPath & "\" & szLocalFile & Chr(34), 0, True)
							Console.WriteLine(" ")
							Console.WriteLine("Installation Success!")
							iResult = 0
						Else
							iResult = 2
						End If
                    End If
                Else
                    iResult = 1
                End If
            Else
                iResult = 1
            End If
        Else
            iResult = 1
        End If
		Console.WriteLine(" ")
        InstallBatFile = iResult
    End Function

    ' *****************************************************************************
    '
    ' FUNCTION INSTALLWINDOWSUPDATE
    '
    ' This Function Install Windows Update File.
    ' Parameters Loads from INI File
    '
    ' PARAMETERS:	szSection -- the Update File Section
    '				szSectionInstall -- the INI File Section
    '
    ' RETURNS:		0 -- the File is Installed With Success
    '				1 -- Some Errors Occur
	'				2 -- Execution Type Mismatch
    '
    ' DEPENDS:		Constants.vb, Constants.sysupdate.vb
    '
    ' ****************************************************************************
    Function InstallWindowsUpdate( szSection As String, szSectionInstall As String) As Integer
        Dim WshShell, fso As Object
        Dim iResult As Integer
        Dim szRootDir As String
        Dim szMainINI As String
        Dim szINI As String
        szRootDir = ROOT_DIRNAME & "\" & FIRMNAME & "\" & PROJECTNAME
		Console.WriteLine("Program Root Directory = {0}", szRootDir)
        szMainINI = szRootDir & "\" & INI_MAINC
		Console.WriteLine( "Main INI File = {0}", szMainINI )
        fso = CreateObject( "Scripting.FileSystemObject" )
        If fso.FileExists(szMainINI) Then
            Dim MainIniFile As New Ini.IniFile(szMainINI)
            szINI = MainIniFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
            szINI = szRootDir & "\" & szINI
			Console.WriteLine("INI FIle = {0}", szINI)
            If fso.FileExists(szINI) Then
                Dim iniFile As New Ini.IniFile(szINI)
                Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
				Dim szBuild As String
                szPath = szRootDir & "\" & DISTRIB_ROOT_FOLDER
                szHost = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                szPref = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                szPort = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                szUser = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                szPass = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                szFileName = iniFile.ReadKey(szSection, "filename")
                szExtension = iniFile.ReadKey(szSection, "extension")
                If Not StrComp(szUser, "", 0) Or Not StrComp(szPref, "http", 0) Then
                    szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                    szLocalFile = szFileName & szExtension
					Console.WriteLine("URL Path = {0}", szURL )
					Console.WriteLine("Local Path = {0}", szPath )
					Console.WriteLine("File: {0}", szLocalFile)
                    If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                        iResult = 1
                    Else
						szBuild = iniFile.ReadKey( szSection, BUILD_PROGRAM )
						If StrComp(szExtension, ".msu", 0) = 0 And StrComp(szBuild, BUILD_MSU_INSTSHELL, 0) = 0 Then
							Console.WriteLine("Build Program = {0}", BUILD_MSU_INSTSHELL )
							WshShell = CreateObject("WScript.Shell")
							MsgBox("wusa.exe " & Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /quiet /norestart")
							'WshShell.Run("wusa.exe " & Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /quiet /norestart", 0, True)
							Console.WriteLine(" ")
							Console.WriteLine("Installation Success!")
							iResult = 0
						Else
							iResult = 2
						End If
                    End If
                Else
                    iResult = 1
                End If
            Else
                iResult = 1
            End If
        Else
            iResult = 1
        End If
		Console.WriteLine(" ")
        InstallWindowsUpdate = iResult
    End Function

    ' *****************************************************************************
    '
    ' FUNCTION INSTALLMSIPROGRAM
    '
    ' This Function Install Windows Update File.
    ' Parameters Loads from INI File
    '
    ' PARAMETERS:	szSection -- the Update File Section
    '				szSectionInstall -- the INI File Section
    '
    ' RETURNS:		0 -- the File is Installed With Success
    '				1 -- Some Errors Occur
	'				2 -- Execution Type Mismatch
	'				3 -- Product Name Not Describe
    '
    ' DEPENDS:		Constants.vb, Constants.sysupdate.vb
    '
    ' ****************************************************************************
    Function InstallMSIProgram( szSection As String, szSectionInstall As String) As Integer
        Dim WshShell, fso As Object
        Dim iResult As Integer
        Dim szRootDir As String
        Dim szMainINI As String
        Dim szINI As String
        szRootDir = ROOT_DIRNAME & "\" & FIRMNAME & "\" & PROJECTNAME
		Console.WriteLine("Program Root Directory = {0}", szRootDir)
        szMainINI = szRootDir & "\" & INI_MAINC
		Console.WriteLine( "Main INI File = {0}", szMainINI )
        fso = CreateObject( "Scripting.FileSystemObject" )
        If fso.FileExists(szMainINI) Then
            Dim MainIniFile As New Ini.IniFile(szMainINI)
            szINI = MainIniFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
            szINI = szRootDir & "\" & szINI
			Console.WriteLine("INI FIle = {0}", szINI)
            If fso.FileExists(szINI) Then
                Dim iniFile As New Ini.IniFile(szINI)
                Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
				Dim szBuild, szProductName As String
                szPath = szRootDir & "\" & DISTRIB_ROOT_FOLDER
                szHost = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                szPref = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                szPort = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                szUser = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                szPass = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                szFileName = iniFile.ReadKey(szSection, "filename")
                szExtension = iniFile.ReadKey(szSection, "extension")
				szProductName = iniFile.ReadKey( szSection, NAME_MSI_PACKETS )
				If StrComp( szProductName, "" ) = 0 Or IsEmpty( szProductName ) Or IsNull( szProductName ) Then
					iResult = 3
				Else
					If Not StrComp(szUser, "", 0) Or Not StrComp(szPref, "http", 0) Then
						szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
						szLocalFile = szFileName & szExtension
						Console.WriteLine("URL Path = {0}", szURL )
						Console.WriteLine("Local Path = {0}", szPath )
						Console.WriteLine("File: {0}", szLocalFile)
						Console.WriteLine("Product Name: {0}", szProductName)
						If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
							iResult = 1
						Else
							szBuild = iniFile.ReadKey( szSection, BUILD_PROGRAM )
							If StrComp(szExtension, ".msi", 0) = 0 And StrComp(szBuild, BUILD_MSI_INSTSHELL, 0) = 0 Then
								Console.WriteLine("Build Program = {0}", BUILD_MSI_INSTSHELL )
								WshShell = CreateObject("WScript.Shell")
								Console.WriteLine("Uninstall {0} Package...", szProductName )
								MsgBox( "wmic.exe product where name=" & Chr(34) & szProductName & Chr(34) & " call uninstall"
								'WshShell.Run("wmic.exe product where name=" & Chr(34) & szProductName & Chr(34) & " call uninstall", 0, True)
								Console.WriteLine("Install {0} Package...", szProductName )
								MsgBox("msiexec /i " & Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /norestart /QN /L*V " & Chr(34) & szPath & "\" & szLocalFile & ".log" & Chr(34))
								'WshShell.Run("msiexec /i " & Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /norestart /QN /L*V " & Chr(34) & szPath & "\" & szLocalFile & ".log" & Chr(34), 0, True)
								Console.WriteLine(" ")
								Console.WriteLine("Installation Success!")
								iResult = 0
							Else
								iResult = 2
							End If
						End If
					End If
                Else
                    iResult = 1
                End If
            Else
                iResult = 1
            End If
        Else
            iResult = 1
        End If
		Console.WriteLine(" ")
        InstallMSIProgram = iResult
    End Function

    ' *****************************************************************************
    '
    ' FUNCTION INSTALL_EXEINNOSETUP
    '
    ' This Function Install Windows Update File.
    ' Parameters Loads from INI File
    '
    ' PARAMETERS:	szSection -- the Update File Section
    '				szSectionInstall -- the INI File Section
    '
    ' RETURNS:		0 -- the File is Installed With Success
    '				1 -- Some Errors Occur
	'				2 -- Execution Type Mismatch
    '
    ' DEPENDS:		Constants.vb, Constants.sysupdate.vb
    '
    ' ****************************************************************************
    Function Install_ExeInnoSetup( szSection As String, szSectionInstall As String) As Integer
        Dim WshShell, fso As Object
        Dim iResult As Integer
        Dim szRootDir As String
        Dim szMainINI As String
        Dim szINI As String
        szRootDir = ROOT_DIRNAME & "\" & FIRMNAME & "\" & PROJECTNAME
		Console.WriteLine("Program Root Directory = {0}", szRootDir)
        szMainINI = szRootDir & "\" & INI_MAINC
		Console.WriteLine( "Main INI File = {0}", szMainINI )
        fso = CreateObject( "Scripting.FileSystemObject" )
        If fso.FileExists(szMainINI) Then
            Dim MainIniFile As New Ini.IniFile(szMainINI)
            szINI = MainIniFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
            szINI = szRootDir & "\" & szINI
			Console.WriteLine("INI FIle = {0}", szINI)
            If fso.FileExists(szINI) Then
                Dim iniFile As New Ini.IniFile(szINI)
                Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
				Dim szBuild As String
                szPath = szRootDir & "\" & DISTRIB_ROOT_FOLDER
                szHost = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                szPref = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                szPort = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                szUser = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                szPass = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                szFileName = iniFile.ReadKey(szSection, "filename")
                szExtension = iniFile.ReadKey(szSection, "extension")
                If Not StrComp(szUser, "", 0) Or Not StrComp(szPref, "http", 0) Then
                    szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                    szLocalFile = szFileName & szExtension
					Console.WriteLine("URL Path = {0}", szURL )
					Console.WriteLine("Local Path = {0}", szPath )
					Console.WriteLine("File: {0}", szLocalFile)
                    If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                        iResult = 1
                    Else
						szBuild = iniFile.ReadKey( szSection, BUILD_PROGRAM )
						If StrComp(szExtension, ".exe", 0) = 0 And StrComp(szBuild, BUILD_EXE_INNOSETUP, 0) = 0 Then
							Console.WriteLine("Build Program = {0}", BUILD_EXE_INNOSETUP)
							WshShell = CreateObject("WScript.Shell")
							Console.WriteLine("Install {0} Package...", szLocalFile )
							MsgBox(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /VERYSILENT /NOCANCEL")
							'WshShell.Run(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /VERYSILENT /NOCANCEL", 0, True)
							Console.WriteLine(" ")
							Console.WriteLine("Installation Success!")
							iResult = 0
						Else
							iResult = 2
						End If
                    End If
                Else
                    iResult = 1
                End If
            Else
                iResult = 1
            End If
        Else
            iResult = 1
        End If
		Console.WriteLine(" ")
        Install_ExeInnoSetup = iResult
    End Function

    ' *****************************************************************************
    '
    ' FUNCTION INSTALL_EXEINSTSHELL
    '
    ' This Function Install Windows Update File.
    ' Parameters Loads from INI File
    '
    ' PARAMETERS:	szSection -- the Update File Section
    '				szSectionInstall -- the INI File Section
    '
    ' RETURNS:		0 -- the File is Installed With Success
    '				1 -- Some Errors Occur
	'				2 -- Execution Type Mismatch
    '
    ' DEPENDS:		Constants.vb, Constants.sysupdate.vb
    '
    ' ****************************************************************************
    Function Install_ExeInstShell( szSection As String, szSectionInstall As String) As Integer
        Dim WshShell, fso As Object
        Dim iResult As Integer
        Dim szRootDir As String
        Dim szMainINI As String
        Dim szINI As String
        szRootDir = ROOT_DIRNAME & "\" & FIRMNAME & "\" & PROJECTNAME
		Console.WriteLine("Program Root Directory = {0}", szRootDir)
        szMainINI = szRootDir & "\" & INI_MAINC
		Console.WriteLine( "Main INI File = {0}", szMainINI )
        fso = CreateObject( "Scripting.FileSystemObject" )
        If fso.FileExists(szMainINI) Then
            Dim MainIniFile As New Ini.IniFile(szMainINI)
            szINI = MainIniFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
            szINI = szRootDir & "\" & szINI
			Console.WriteLine("INI FIle = {0}", szINI)
            If fso.FileExists(szINI) Then
                Dim iniFile As New Ini.IniFile(szINI)
                Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
				Dim szBuild As String
                szPath = szRootDir & "\" & DISTRIB_ROOT_FOLDER
                szHost = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                szPref = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                szPort = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                szUser = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                szPass = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                szFileName = iniFile.ReadKey(szSection, "filename")
                szExtension = iniFile.ReadKey(szSection, "extension")
                If Not StrComp(szUser, "", 0) Or Not StrComp(szPref, "http", 0) Then
                    szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                    szLocalFile = szFileName & szExtension
					Console.WriteLine("URL Path = {0}", szURL )
					Console.WriteLine("Local Path = {0}", szPath )
					Console.WriteLine("File: {0}", szLocalFile)
                    If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                        iResult = 1
                    Else
						szBuild = iniFile.ReadKey( szSection, BUILD_PROGRAM )
						If StrComp(szExtension, ".exe", 0) = 0 And StrComp(szBuild, BUILD_EXE_INSTSHELL, 0) = 0 Then
							Console.WriteLine("Build Program = {0}", BUILD_EXE_INSTSHELL )
							WshShell = CreateObject("WScript.Shell")
							Console.WriteLine("Install {0} Package...", szLocalFile )
							MsgBox(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /quiet /norestart")
							'WshShell.Run(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /quiet /norestart", 0, True)
							Console.WriteLine(" ")
							Console.WriteLine("Installation Success!")
							iResult = 0
						Else
							iResult = 2
						End If
                    End If
                Else
                    iResult = 1
                End If
            Else
                iResult = 1
            End If
        Else
            iResult = 1
        End If
		Console.WriteLine(" ")
        Install_ExeInstShell = iResult
    End Function

    ' *****************************************************************************
    '
    ' FUNCTION INSTALL_EXEJAVA
    '
    ' This Function Install Windows Update File.
    ' Parameters Loads from INI File
    '
    ' PARAMETERS:	szSection -- the Update File Section
    '				szSectionInstall -- the INI File Section
    '
    ' RETURNS:		0 -- the File is Installed With Success
    '				1 -- Some Errors Occur
	'				2 -- Execution Type Mismatch
    '
    ' DEPENDS:		Constants.vb, Constants.sysupdate.vb
    '
    ' ****************************************************************************
    Function Install_ExeJava( szSection As String, szSectionInstall As String) As Integer
        Dim WshShell, fso As Object
        Dim iResult As Integer
        Dim szRootDir As String
        Dim szMainINI As String
        Dim szINI As String
        szRootDir = ROOT_DIRNAME & "\" & FIRMNAME & "\" & PROJECTNAME
		Console.WriteLine("Program Root Directory = {0}", szRootDir)
        szMainINI = szRootDir & "\" & INI_MAINC
		Console.WriteLine( "Main INI File = {0}", szMainINI )
        fso = CreateObject( "Scripting.FileSystemObject" )
        If fso.FileExists(szMainINI) Then
            Dim MainIniFile As New Ini.IniFile(szMainINI)
            szINI = MainIniFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
            szINI = szRootDir & "\" & szINI
			Console.WriteLine("INI FIle = {0}", szINI)
            If fso.FileExists(szINI) Then
                Dim iniFile As New Ini.IniFile(szINI)
                Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
				Dim szBuild As String
                szPath = szRootDir & "\" & DISTRIB_ROOT_FOLDER
                szHost = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                szPref = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                szPort = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                szUser = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                szPass = MainIniFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                szFileName = iniFile.ReadKey(szSection, "filename")
                szExtension = iniFile.ReadKey(szSection, "extension")
                If Not StrComp(szUser, "", 0) Or Not StrComp(szPref, "http", 0) Then
                    szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                    szLocalFile = szFileName & szExtension
					Console.WriteLine("URL Path = {0}", szURL )
					Console.WriteLine("Local Path = {0}", szPath )
					Console.WriteLine("File: {0}", szLocalFile)
                    If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                        iResult = 1
                    Else
						szBuild = iniFile.ReadKey( szSection, BUILD_PROGRAM )
						If StrComp(szExtension, ".exe", 0) = 0 And StrComp(szBuild, BUILD_EXE_JAVA, 0) = 0 Then
							Console.WriteLine("Build Program = {0}", BUILD_EXE_JAVA )
							WshShell = CreateObject("WScript.Shell")
							Console.WriteLine("Install {0} Package...", szLocalFile )
							MsgBox(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /s")
							'WshShell.Run(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /s", 0, True)
							Console.WriteLine(" ")
							Console.WriteLine("Installation Success!")
							iResult = 0
						Else
							iResult = 2
						End If
                    End If
                Else
                    iResult = 1
                End If
            Else
                iResult = 1
            End If
        Else
            iResult = 1
        End If
		Console.WriteLine(" ")
        Install_ExeJava = iResult
    End Function

    ' ****************************************************************************
    '
    ' FUNCTION Registry_LoadEx
    ' This Function is Load Update Installing Subroutine at Registry Entry
	' Parameters Load From INI Files
    '
    ' PARAMETERS:   strRoot -- the Root of Running Program
	'				szSectionInstall -- Root Section of Installing Program
    ' RETURNS:      0 -- if Success and Continued Process
	'				-1 -- if Success And Terminated Process 
    '               1 -- if strProgram Not Found
    '               2 -- if Main INI File not Found
    '               3 -- if Section INI File not Found
    '				4 -- if Installer Update File Not Found
	'				5 -- if PROGRAMNAME_IN_REGISTRY File Not Found
	'
	' DEPENDS:		Constants.sysupdate.vb
	'
    ' ****************************************************************************
    Function Registry_LoadEx(strRoot As String, szSectionInstall As String) As Integer
        Dim iResult As Integer
        Dim fso, WshShell As Object
        Dim strProgram As String
        fso = CreateObject("Scripting.FileSystemObject ")
            Dim szMainINI As String
            Dim szINI As String
            szMainINI = strRoot & "\" & INI_MAIN
			Console.WriteLine( "Main INI File = {0}", szMainINI )
            If fso.FileExists(szMainINI) Then
                Dim MainIniFile As New Ini.IniFile(szMainINI)
                szINI = MainIniFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = strRoot & "\" & szINI
            	Console.WriteLine( "INI File = {0}", szINI )
				If fso.FileExists(szINI) Then
					Dim strLocal_File1 As String
					Dim iniFile As New Ini.IniFile(szINI)
					strProgram = iniFile.ReadKey( szSectionInstall, REGISTRY_PROGRAM )
					strLocal_File1 = strRoot & "\" & strProgram
					Console.WriteLine( REGISTRY_PROGRAM & " = {0}", strLocal_File1 )
					If fso.FileExists( strLocal_File1 ) Then
						Dim strLocal_File2 As String
						strLocal_File2 = strRoot & "\" & PROGRAMNAME_IN_REGISTRY
						Console.WriteLine( "Program In Registry = {0}", strLocal_File1 )
						If fso.FileExists( strLocal_File2 ) Then
							Const SZREG_TITLE As String = "TITLE"
							Const SZREG_1 As String = "1"
							Const SZREG_2 As String = "2"
							Const SZREG_I As String = "100\"
							Dim szRegValue As String
							WshShell = CreateObject("WScript.Shell")
							WshShell.RegWrite(REGISTRY_KEY, "")
							szRegValue = REGISTRY_TITLE_MAIN
							WshShell.RegWrite(REGISTRY_KEY & "\" & SZREG_TITLE, szRegValue, "REG_SZ" )
							szRegValue = iniFile.ReadKey( szSectionInstall, REGISTRY_KEYTITLE )
							WshShell.RegWrite(REGISTRY_KEY & "\" & SZREG_I, szRegValue, "REG_SZ" )
							szRegValue = strLocal_File1
							WshShell.RegWrite(REGISTRY_KEY & "\" & SZREG_I & SZREG_1, szRegValue, "REG_SZ" )
							szRegValue = iniFile.ReadKey( szSectionInstall, NEXT_SECTIONINSTALL )
							If StrComp( szRegValue, "" ) = 0 Or IsEmpty( szRegValue ) Or IsNull( szRegValue ) Then
								iResult = -1
								Console.WriteLine(" ")
								Console.WriteLine("Ending Installation...")
							Else
								szRegValue = strLocal_File2 & " " & Chr(34) & szRegValue & Chr(34)
								WshShell.RegWrite(REGISTRY_KEY & "\" & SZREG_I & SZREG_2, szRegValue, "REG_SZ" )
								iResult = 0
								Console.WriteLine("Registry Value = {0}", szRegValue)
								Console.WriteLine(" ")
								Console.WriteLine("Continue Installation...")
							End If
						Else
							iResult = 5
						End If
					Else
						iResult = 4
					End If
                Else
                    iResult = 3
                End If
            Else
                iResult = 2
            End If
		WriteLine(" ")
		Registry_LoadEx = iResult
    End Function

End Module
