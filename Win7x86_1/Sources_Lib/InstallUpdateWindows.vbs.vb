Namespace InstallUpdateWindows

    Public Class InstallUpdateWindows

    Public Const INI_MAIN As String = "main.ini"

    Public Const HOSTUPDATE_SECTION As String = "host_update"   ' Main Host Downloaded From Section
    Public Const HTTPHOST_KEY As String = "host"        ' Domain Host Name or IP
    Public Const HTTPPREF_KEY As String = "pref"        ' Host Prefix (http/https/ftp)
    Public Const HTTPPORT_KEY As String = "port"        ' Host port (standart or Non-Standart)
    Public Const HTTPUSER_KEY As String = "user"        ' User Authentication Name (May be Empty)
    Public Const HTTPPASS_KEY As String = "password"    ' User Authentication Password (Empty only if User Absent)
    Public Const MAINFOLDER_KEY As String = "mainfolders"   ' Path to Downloaded from Folder 
    Public Const UPDATE_NET_SECTION As String = "NET Framework"
    Public Const UPDATE_IE11_SECTION1 As String = "IE11 Update"
    Public Const UPDATE_IE11_SECTION2 As String = "IE11"
    Public Const UPDATE_WNF_SECTION As String = "WinMF 5.1"
    Public Const UPDATE_TLS_SECTION As String = "TLS 1.2"
    Public Const UPDATE_SEC_SECTION As String = "SecUpdate 01"
    Public Const UPDATE_OTHER_SECTIUON As String = "Other Update"
    Public Const DISTRIB_FOLDER_KEY As String = "folder"
    Public Const SOME_INIFILE_KEY As String = "ini_file"
    Public Const REGISTRY_KEYTITLE As String = "keytitle"
    Public Const REGISTRY_PROGRAM As String = "regprogram"
	Public Const NEXT_SECTIONINSTALL As String = "next_sectioninstall"
	Public Const BUILD_PROGRAM As String = "build"
	Public Const NAME_MSI_PACKETS As String = "productname"
	Public Const INITIAL_SECTION As String = "ini_section"

#Region "Private Data"

		Private Dim strProjectRoot_IUW As String
		Private Dim strFileMainINI_IUM As String
        Private MainINIFile As New Ini.IniFile
        Private Dim iResult_IUM As Integer

#End Region '/Private Data

#Region " ŒÕ—“–” “Œ–"

		Public Sub New( ByVal strProjectRoot As String )
			Dim fso As Object
			fso = CreateObject("Scripting.FileSystemObject")
			strProjectRoot_IUW = strProjectRoot
			Console.WriteLine( "Main Root Folder = {0}", strProjectRoot_IUW )
			If Not fso.FolderExists( strProjectRoot_IUW ) Then
				iResult_IUM = 1
			Else
				strFileMainINI_IUM = strProjectRoot_IUW & "\" & INI_MAIN
				Console.WriteLine( "Main INI File = {0}", strFileMainINI_IUM )
				if Not fso.FileExists( strFileMainINI_IUM ) Then
					iResult_IUM = 2
				Else
                    MainINIFile.SetIniPath(strFileMainINI_IUM)
                    iResult_IUM = 0
				End If
			End if			
		End Sub
#End Region '/ ŒÕ—“–” “Œ–

#Region "Ã≈“Œƒ€ –¿¡Œ“€ — INI-‘¿…À¿Ã»"

        Public Function Is_Result() As Integer
            Is_Result = iResult_IUM
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
        ' RETURNS:      0 -- if Success and Continued Process
        '				-1 -- if Success And Terminated Process 
        '               1 -- if strProjectRoot Not Found
        '               2 -- if Main INI File not Found
        '               3 -- if Section INI File not Found
        '				4 -- if Installer Update File Not Found
        '				5 -- if PROGRAMNAME_IN_REGISTRY File Not Found
        '				6 -- if Download And Save File Error
        '				7 -- Execution Type Mismatch
        '				100 -- if Some Unknown Error Occur
        '
        ' DEPENDS:		Constants.vb, Constants.sysupdate.vb
        '
        ' ****************************************************************************
        Public Sub InstallBatFile( szSection As String, szSectionInstall As String)
        Dim WshShell, fso As Object
        Dim szINI As String
            If iResult_IUM > 0 Then
                fso = CreateObject("Scripting.FileSystemObject")
                szINI = MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If fso.FileExists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
                    szHost = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                    szPref = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                    szPort = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                    szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                    szUser = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                    szPass = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                    szFileName = iniFile.ReadKey(szSection, "filename")
                    szExtension = iniFile.ReadKey(szSection, "extension")
                    If Not StrComp(szUser, "", 0) Or Not StrComp(szPref, "http", 0) Then
                        szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                        szLocalFile = szFileName & szExtension
                        Console.WriteLine("URL Path = {0}", szURL)
                        Console.WriteLine("Local Path = {0}", szPath)
                        Console.WriteLine("File: {0}", szLocalFile)
                        If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                            iResult_IUM = 6
                        Else
                            szBuild = iniFile.ReadKey(szSection, BUILD_PROGRAM)
                            If StrComp(szExtension, ".bat", 0) = 0 Or StrComp(szExtension, ".cmd", 0) = 0 And StrComp(szBuild, BUILD_BAT, 0) = 0 Then
                                Console.WriteLine("Build Program = {0}", BUILD_BAT)
                                WshShell = CreateObject("WScript.Shell")
                                MsgBox("cmd /c cd /d " & Chr(34) & szPath & Chr(34) & " && " & Chr(34) & szPath & "\" & szLocalFile & Chr(34))
                                'WshShell.Run("cmd /c cd /d " & Chr(34) & szPath & Chr(34) & " && " & Chr(34) & szPath & "\" & szLocalFile & Chr(34), 0, True)
                                Console.WriteLine(" ")
                                Console.WriteLine("Installation Success!")
                                iResult_IUM = 0
                            Else
                                iResult_IUM = 7
                            End If
                        End If
                    Else
                        iResult_IUM = 100
                    End If
                Else
                    iResult_IUM = 3
                End If
            End If
            Console.WriteLine(" ")
    End Sub


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
    ' DEPENDS:		Constants.vb, Constants.sysupdate.vb
    '
    ' ****************************************************************************
    Public Sub InstallWindowsUpdate( szSection As String, szSectionInstall As String)
            Dim WshShell, fso As Object
            Dim szINI As String
            If iResult_IUM > 0 Then
                fso = CreateObject("Scripting.FileSystemObject")
                szINI = MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If fso.FileExists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
                    szHost = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                    szPref = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                    szPort = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                    szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                    szUser = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                    szPass = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                    szFileName = iniFile.ReadKey(szSection, "filename")
                    szExtension = iniFile.ReadKey(szSection, "extension")
                    If Not StrComp(szUser, "", 0) Or Not StrComp(szPref, "http", 0) Then
                        szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                        szLocalFile = szFileName & szExtension
                        Console.WriteLine("URL Path = {0}", szURL)
                        Console.WriteLine("Local Path = {0}", szPath)
                        Console.WriteLine("File: {0}", szLocalFile)
                        If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                            iResult_IUM = 6
                        Else
                            szBuild = iniFile.ReadKey(szSection, BUILD_PROGRAM)
                            If StrComp(szExtension, ".msu", 0) = 0 And StrComp(szBuild, BUILD_MSU_INSTSHELL, 0) = 0 Then
                                Console.WriteLine("Build Program = {0}", BUILD_MSU_INSTSHELL)
                                WshShell = CreateObject("WScript.Shell")
                                MsgBox("wusa.exe " & Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /quiet /norestart")
                                'WshShell.Run("wusa.exe " & Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /quiet /norestart", 0, True)
                                Console.WriteLine(" ")
                                Console.WriteLine("Installation Success!")
                                iResult_IUM = 0
                            Else
                                iResult_IUM = 7
                            End If
                        End If
                    Else
                        iResult_IUM = 100
                    End If
                Else
                    iResult_IUM = 3
                End If
            End If
            Console.WriteLine(" ")
    End Sub

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
    Public Sub InstallMSIProgram( szSection As String, szSectionInstall As String)
            Dim WshShell, fso As Object
            Dim szINI As String
            If iResult_IUM > 0 Then
                fso = CreateObject("Scripting.FileSystemObject")
                szINI = MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If fso.FileExists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
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
                        Console.WriteLine("URL Path = {0}", szURL)
                        Console.WriteLine("Local Path = {0}", szPath)
                        Console.WriteLine("File: {0}", szLocalFile)
                        If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                            Me.iResult_IUM = 6
                        Else
                            Dim szProductName As String
                            szBuild = iniFile.ReadKey(szSection, BUILD_PROGRAM)
                            szProductName = iniFile.ReadKey(szSection, NAME_MSI_PACKETS)
                            If StrComp(szProductName, "") <> 0 And Not szProductName = Nothing Then
                                If StrComp(szExtension, ".msi", 0) = 0 And StrComp(szBuild, BUILD_MSI_INSTSHELL, 0) = 0 Then
                                    Console.WriteLine("Build Program = {0}", BUILD_MSI_INSTSHELL)
                                    WshShell = CreateObject("WScript.Shell")
                                    Console.WriteLine("Uninstall {0} Package...", szProductName)
                                    MsgBox("wmic.exe product where name=" & Chr(34) & szProductName & Chr(34) & " call uninstall")
                                    'WshShell.Run("wmic.exe product where name=" & Chr(34) & szProductName & Chr(34) & " call uninstall", 0, True)
                                    Console.WriteLine("Install {0} Package...", szProductName)
                                    MsgBox("msiexec /i " & Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /norestart /QN /L*V " & Chr(34) & szPath & "\" & szLocalFile & ".log" & Chr(34))
                                    'WshShell.Run("msiexec /i " & Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /norestart /QN /L*V " & Chr(34) & szPath & "\" & szLocalFile & ".log" & Chr(34), 0, True)
                                    Console.WriteLine(" ")
                                    Console.WriteLine("Installation Success!")
                                    Me.iResult_IUM = 0
                                Else
                                    Me.iResult_IUM = 7
                                End If
                            Else
                                Me.iResult_IUM = 3
                            End If
                        End If
                    Else
                        Me.iResult_IUM = 100
                    End If
                Else
                    Me.iResult_IUM = 3
                End If
            End If
            Console.WriteLine(" ")
    End Sub

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
    Public Sub Install_ExeInnoSetup( szSection As String, szSectionInstall As String)
            Dim WshShell, fso As Object
            Dim szINI As String
            If iResult_IUM > 0 Then
                fso = CreateObject("Scripting.FileSystemObject")
                szINI = MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If fso.FileExists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
                    szHost = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                    szPref = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                    szPort = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                    szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                    szUser = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                    szPass = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                    szFileName = iniFile.ReadKey(szSection, "filename")
                    szExtension = iniFile.ReadKey(szSection, "extension")
                    If Not StrComp(szUser, "", 0) Or Not StrComp(szPref, "http", 0) Then
                        szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                        szLocalFile = szFileName & szExtension
                        Console.WriteLine("URL Path = {0}", szURL)
                        Console.WriteLine("Local Path = {0}", szPath)
                        Console.WriteLine("File: {0}", szLocalFile)
                        If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                            iResult_IUM = 6
                        Else
                            szBuild = iniFile.ReadKey(szSection, BUILD_PROGRAM)
                            If StrComp(szExtension, ".exe", 0) = 0 And StrComp(szBuild, BUILD_EXE_INNOSETUP, 0) = 0 Then
                                Console.WriteLine("Build Program = {0}", BUILD_EXE_INNOSETUP)
                                WshShell = CreateObject("WScript.Shell")
                                Console.WriteLine("Install {0} Package...", szLocalFile)
                                MsgBox(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /VERYSILENT /NOCANCEL")
                                'WshShell.Run(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /VERYSILENT /NOCANCEL", 0, True)
                                Console.WriteLine(" ")
                                Console.WriteLine("Installation Success!")
                                iResult_IUM = 0
                            Else
                                iResult_IUM = 7
                            End If
                        End If
                    Else
                        iResult_IUM = 100
                    End If
                Else
                    iResult_IUM = 3
                End If
            End If
            Console.WriteLine(" ")
    End Sub

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
    Public Sub Install_ExeInstShell( szSection As String, szSectionInstall As String)
            Dim WshShell, fso As Object
            Dim szINI As String
            If iResult_IUM > 0 Then
                fso = CreateObject("Scripting.FileSystemObject")
                szINI = MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If fso.FileExists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
                    szHost = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                    szPref = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                    szPort = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                    szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                    szUser = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                    szPass = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                    szFileName = iniFile.ReadKey(szSection, "filename")
                    szExtension = iniFile.ReadKey(szSection, "extension")
                    If Not StrComp(szUser, "", 0) Or Not StrComp(szPref, "http", 0) Then
                        szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                        szLocalFile = szFileName & szExtension
                        Console.WriteLine("URL Path = {0}", szURL)
                        Console.WriteLine("Local Path = {0}", szPath)
                        Console.WriteLine("File: {0}", szLocalFile)
                        If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                            iResult_IUM = 6
                        Else
                            szBuild = iniFile.ReadKey(szSection, BUILD_PROGRAM)
                            If StrComp(szExtension, ".exe", 0) = 0 And StrComp(szBuild, BUILD_EXE_INSTSHELL, 0) = 0 Then
                                Console.WriteLine("Build Program = {0}", BUILD_EXE_INSTSHELL)
                                WshShell = CreateObject("WScript.Shell")
                                Console.WriteLine("Install {0} Package...", szLocalFile)
                                MsgBox(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /quiet /norestart")
                                'WshShell.Run(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /quiet /norestart", 0, True)
                                Console.WriteLine(" ")
                                Console.WriteLine("Installation Success!")
                                iResult_IUM = 0
                            Else
                                iResult_IUM = 7
                            End If
                        End If
                    Else
                        iResult_IUM = 100
                    End If
                Else
                    iResult_IUM = 3
                End If
            End If
            Console.WriteLine(" ")
    End Sub

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
    Public Sub Install_ExeJava( szSection As String, szSectionInstall As String)
            Dim WshShell, fso As Object
            Dim szINI As String
            If iResult_IUM > 0 Then
                fso = CreateObject("Scripting.FileSystemObject")
                szINI = MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If fso.FileExists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
                    szHost = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                    szPref = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                    szPort = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                    szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                    szUser = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                    szPass = MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                    szFileName = iniFile.ReadKey(szSection, "filename")
                    szExtension = iniFile.ReadKey(szSection, "extension")
                    If Not StrComp(szUser, "", 0) Or Not StrComp(szPref, "http", 0) Then
                        szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                        szLocalFile = szFileName & szExtension
                        Console.WriteLine("URL Path = {0}", szURL)
                        Console.WriteLine("Local Path = {0}", szPath)
                        Console.WriteLine("File: {0}", szLocalFile)
                        If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                            iResult_IUM = 6
                        Else
                            szBuild = iniFile.ReadKey(szSection, BUILD_PROGRAM)
                            If StrComp(szExtension, ".exe", 0) = 0 And StrComp(szBuild, BUILD_EXE_JAVA, 0) = 0 Then
                                Console.WriteLine("Build Program = {0}", BUILD_EXE_JAVA)
                                WshShell = CreateObject("WScript.Shell")
                                Console.WriteLine("Install {0} Package...", szLocalFile)
                                MsgBox(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /s")
                                'WshShell.Run(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /s", 0, True)
                                Console.WriteLine(" ")
                                Console.WriteLine("Installation Success!")
                                iResult_IUM = 0
                            Else
                                iResult_IUM = 7
                            End If
                        End If
                    Else
                        iResult_IUM = 100
                    End If
                Else
                    iResult_IUM = 3
                End If
            End If
            Console.WriteLine(" ")
    End Sub

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

#End Region '/Ã≈“Œƒ€ –¿¡Œ“€ — INI-‘¿…À¿Ã»

    End Class

End Namespace


