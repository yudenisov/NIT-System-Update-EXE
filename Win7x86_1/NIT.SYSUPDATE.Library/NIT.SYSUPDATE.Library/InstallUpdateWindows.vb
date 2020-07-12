Option Strict On
Imports System
Imports Microsoft.Win32
Imports System.IO
Imports System.Security
Imports System.Net
Imports System.Text
Imports Microsoft.VisualBasic.Devices

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

        Private strProjectRoot_IUW As String
        Private strFileMainINI_IUM As String
        Private MainINIFile As New Ini.IniFile
        Private iResult_IUM As Integer

#End Region '/Private Data

#Region "КОНСТРУКТОР"

        Public Sub New(ByVal strProjectRoot As String)
            Me.strProjectRoot_IUW = strProjectRoot
            Console.WriteLine("Main Root Folder = {0}", Me.strProjectRoot_IUW)
            If Not IO.Directory.Exists(Me.strProjectRoot_IUW) Then
                Me.iResult_IUM = 1
            Else
                Me.strFileMainINI_IUM = Me.strProjectRoot_IUW & "\" & INI_MAIN
                Console.WriteLine("Main INI File = {0}", Me.strFileMainINI_IUM)
                If Not IO.File.Exists(Me.strFileMainINI_IUM) Then
                    Me.iResult_IUM = 2
                Else
                    Me.MainINIFile.SetIniPath(Me.strFileMainINI_IUM)
                    Me.iResult_IUM = 0
                End If
            End If
        End Sub

        Public Sub Set_Root(ByVal strProjectRoot As String)
            Dim fso As Object
            fso = CreateObject("Scripting.FileSystemObject")
            Me.strProjectRoot_IUW = strProjectRoot
            Console.WriteLine("Main Root Folder = {0}", Me.strProjectRoot_IUW)
            If Not IO.Directory.Exists(Me.strProjectRoot_IUW) Then
                Me.iResult_IUM = 1
            Else
                Me.strFileMainINI_IUM = Me.strProjectRoot_IUW & "\" & INI_MAIN
                Console.WriteLine("Main INI File = {0}", Me.strFileMainINI_IUM)
                If Not IO.File.Exists(Me.strFileMainINI_IUM) Then
                    Me.iResult_IUM = 2
                Else
                    Me.MainINIFile.SetIniPath(Me.strFileMainINI_IUM)
                    Me.iResult_IUM = 0
                End If
            End If
        End Sub

        Public Sub New()
            Dim strProjectRoot As String
            Dim fso As Object
            strProjectRoot = ROOT_DIRNAME & "\" & FIRMNAME & "\" & PROJECTNAME
            fso = CreateObject("Scripting.FileSystemObject")
            Me.strProjectRoot_IUW = strProjectRoot
            Console.WriteLine("Main Root Folder = {0}", Me.strProjectRoot_IUW)
            If Not IO.Directory.Exists(Me.strProjectRoot_IUW) Then
                Me.iResult_IUM = 1
            Else
                Me.strFileMainINI_IUM = Me.strProjectRoot_IUW & "\" & INI_MAIN
                Console.WriteLine("Main INI File = {0}", Me.strFileMainINI_IUM)
                If Not IO.File.Exists(Me.strFileMainINI_IUM) Then
                    Me.iResult_IUM = 2
                Else
                    Me.MainINIFile.SetIniPath(Me.strFileMainINI_IUM)
                    Me.iResult_IUM = 0
                End If
            End If
        End Sub
#End Region '/КОНСТРУКТОР

#Region "МЕТОДЫ РАБОТЫ С INI-ФАЙЛАМИ"

        Public Function Is_Result() As Integer
            Is_Result = Me.iResult_IUM
        End Function

        Public Function Is_IResult_0(iResult As Integer) As Integer
            If iResult > 0 Then
                Is_IResult_0 = iResult
            Else
                Is_IResult_0 = iResult_IUM
            End If
            iResult_IUM = 0
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
        Public Sub InstallBatFile(ByVal szSection As String, ByVal szSectionInstall As String)
            Dim szINI As String
            If Me.iResult_IUM <= 0 Then
                szINI = Me.MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If IO.File.Exists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
                    szHost = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                    szPref = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                    szPort = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                    szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                    szUser = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                    szPass = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                    szFileName = iniFile.ReadKey(szSection, "filename")
                    szExtension = iniFile.ReadKey(szSection, "extension")
                    If szUser = Nothing Or szUser.Length = 0 Then
                        If StrComp(szPref, "http", vbTextCompare) = 0 Then
							szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                            szLocalFile = szFileName & szExtension
                            Console.WriteLine("URL Path = {0}", szURL)
                            Console.WriteLine("Local Path = {0}", szPath)
                            Console.WriteLine("File: {0}", szLocalFile)
                            If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                                Me.iResult_IUM = 6
                            Else
                                szBuild = iniFile.ReadKey(szSection, BUILD_PROGRAM)
                                If (StrComp(szExtension, ".bat", 0) = 0 Or StrComp(szExtension, ".cmd", 0) = 0) And StrComp(szBuild, BUILD_BAT, 0) = 0 Then
                                    Console.WriteLine("Build Program = {0}", BUILD_BAT)
                                    'MsgBox("C:\Windows\System32\cmd.exe /c " & Chr(34) & szPath & "\" & szLocalFile & Chr(34))
                                    ' Attention!!!
                                    Dim myProcess As Process
                                    myProcess = System.Diagnostics.Process.Start("C:\Windows\System32\cmd.exe", "/c " & Chr(34) & szPath & "\" & szLocalFile & Chr(34))
                                    myProcess.WaitForExit()
                                    myProcess.Close()
                                    Console.WriteLine(" ")
                                    Console.WriteLine("Installation Success!")
                                    Me.iResult_IUM = 0
                                Else
                                    Me.iResult_IUM = 7
                            End If
                        End If
                        Else
                            Me.iResult_IUM = 100
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
        Public Sub InstallWindowsUpdate(ByVal szSection As String, ByVal szSectionInstall As String)
            Dim szINI As String
            If Me.iResult_IUM <= 0 Then
                szINI = Me.MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If IO.File.Exists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
                    szHost = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                    szPref = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                    szPort = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                    szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                    szUser = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                    szPass = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                    szFileName = iniFile.ReadKey(szSection, "filename")
                    szExtension = iniFile.ReadKey(szSection, "extension")
                    If szUser = Nothing Or szUser.Length = 0 Then
                        If StrComp(szPref, "http", vbTextCompare) = 0 Then
                            szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                            szLocalFile = szFileName & szExtension
                            Console.WriteLine("URL Path = {0}", szURL)
                            Console.WriteLine("Local Path = {0}", szPath)
                            Console.WriteLine("File: {0}", szLocalFile)
                            If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                                Me.iResult_IUM = 6
                            Else
                                szBuild = iniFile.ReadKey(szSection, BUILD_PROGRAM)
                                If StrComp(szExtension, ".msu", 0) = 0 And StrComp(szBuild, BUILD_MSU_INSTSHELL, 0) = 0 Then
                                    Console.WriteLine("Build Program = {0}", BUILD_MSU_INSTSHELL)
                                    'MsgBox("C:\Windows\System32\wusa.exe " & Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /quiet /norestart")
                                    ' Attention!!!
                                    Dim myProcess As Process
                                    myProcess = System.Diagnostics.Process.Start("C:\Windows\System32\wusa.exe", Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /quiet /norestart")
                                    myProcess.WaitForExit()
                                    myProcess.Close()
                                    Console.WriteLine(" ")
                                    Console.WriteLine("Installation Success!")
                                    Me.iResult_IUM = 0
                                Else
                                    Me.iResult_IUM = 7
                                End If
                            End If
                        Else
                            Me.iResult_IUM = 100
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
        Public Sub InstallMSIProgram(ByVal szSection As String, ByVal szSectionInstall As String)
            Dim szINI As String
            If Me.iResult_IUM <= 0 Then
                szINI = Me.MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If IO.File.Exists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
                    szHost = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                    szPref = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                    szPort = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                    szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                    szUser = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                    szPass = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                    szFileName = iniFile.ReadKey(szSection, "filename")
                    szExtension = iniFile.ReadKey(szSection, "extension")
                    If szUser = Nothing Or szUser.Length = 0 Then
                        If StrComp(szPref, "http", vbTextCompare) = 0 Then
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
                                        Console.WriteLine("Uninstall {0} Package...", szProductName)
                                        'MsgBox("C:\Windows\System32\wbem\WMIC.exe product where name=" & Chr(34) & szProductName & Chr(34) & " call uninstall")
                                        ' Attention!!!
                                        Dim myProcess As Process
                                        myProcess = System.Diagnostics.Process.Start("C:\Windows\System32\wbem\WMIC.exe", " product where name=" & Chr(34) & szProductName & Chr(34) & " call uninstall")
                                        myProcess.WaitForExit()
                                        myProcess.Close()
                                        Console.WriteLine("Install {0} Package...", szProductName)
                                        'MsgBox("C:\Windows\System32\msiexec.exe /i " & Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /norestart /QN /L*V " & Chr(34) & szPath & "\" & szLocalFile & ".log" & Chr(34))
                                        ' Attention!!!
                                        myProcess = System.Diagnostics.Process.Start("C:\Windows\System32\msiexec.exe", " /i " & Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /norestart /QN /L*V " & Chr(34) & szPath & "\" & szLocalFile & ".log" & Chr(34))
                                        myProcess.WaitForExit()
                                        myProcess.Close()
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
        Public Sub Install_ExeInnoSetup(ByVal szSection As String, ByVal szSectionInstall As String)
            Dim szINI As String
            If Me.iResult_IUM <= 0 Then
                szINI = Me.MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If IO.File.Exists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
                    szHost = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                    szPref = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                    szPort = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                    szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                    szUser = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                    szPass = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                    szFileName = iniFile.ReadKey(szSection, "filename")
                    szExtension = iniFile.ReadKey(szSection, "extension")
                    If szUser = Nothing Or szUser.Length = 0 Then
                        If StrComp(szPref, "http", vbTextCompare) = 0 Then
                            szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                            szLocalFile = szFileName & szExtension
                            Console.WriteLine("URL Path = {0}", szURL)
                            Console.WriteLine("Local Path = {0}", szPath)
                            Console.WriteLine("File: {0}", szLocalFile)
                            If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                                Me.iResult_IUM = 6
                            Else
                                szBuild = iniFile.ReadKey(szSection, BUILD_PROGRAM)
                                If StrComp(szExtension, ".exe", 0) = 0 And StrComp(szBuild, BUILD_EXE_INNOSETUP, 0) = 0 Then
                                    Console.WriteLine("Build Program = {0}", BUILD_EXE_INNOSETUP)
                                    Console.WriteLine("Install {0} Package...", szLocalFile)
                                    'MsgBox(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /VERYSILENT /NOCANCEL")
                                    ' Attention!!!
                                    Dim myProcess As Process
                                    myProcess = System.Diagnostics.Process.Start(Chr(34) & szPath & "\" & szLocalFile & Chr(34), " /VERYSILENT /NOCANCEL")
                                    myProcess.WaitForExit()
                                    myProcess.Close()
                                    Console.WriteLine(" ")
                                    Console.WriteLine("Installation Success!")
                                    Me.iResult_IUM = 0
                                Else
                                    Me.iResult_IUM = 7
                                End If
                            End If
                        Else
                            Me.iResult_IUM = 100
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
        Public Sub Install_ExeInstShell(ByVal szSection As String, ByVal szSectionInstall As String)
            Dim szINI As String
            If Me.iResult_IUM <= 0 Then
                szINI = Me.MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If IO.File.Exists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
                    szHost = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                    szPref = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                    szPort = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                    szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                    szUser = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                    szPass = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                    szFileName = iniFile.ReadKey(szSection, "filename")
                    szExtension = iniFile.ReadKey(szSection, "extension")
                    If szUser = Nothing Or szUser.Length = 0 Then
                        If StrComp(szPref, "http", vbTextCompare) = 0 Then
                            szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                            szLocalFile = szFileName & szExtension
                            Console.WriteLine("URL Path = {0}", szURL)
                            Console.WriteLine("Local Path = {0}", szPath)
                            Console.WriteLine("File: {0}", szLocalFile)
                            If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                                Me.iResult_IUM = 6
                            Else
                                szBuild = iniFile.ReadKey(szSection, BUILD_PROGRAM)
                                If StrComp(szExtension, ".exe", vbTextCompare) = 0 And StrComp(szBuild, BUILD_EXE_INSTSHELL, vbBinaryCompare) = 0 Then
                                    Console.WriteLine("Build Program = {0}", BUILD_EXE_INSTSHELL)
                                    Console.WriteLine("Install {0} Package...", szLocalFile)
                                    'MsgBox(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /quiet /norestart")
                                    ' Attention!!!
                                    Dim myProcess As Process
                                    myProcess = System.Diagnostics.Process.Start(Chr(34) & szPath & "\" & szLocalFile & Chr(34), " /quiet /norestart")
                                    myProcess.WaitForExit()
                                    myProcess.Close()
                                    Console.WriteLine(" ")
                                    Console.WriteLine("Installation Success!")
                                    Me.iResult_IUM = 0
                                Else
                                    Me.iResult_IUM = 7
                                End If
                            End If
                        Else
                            Me.iResult_IUM = 100
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
        Public Sub Install_ExeJava(ByVal szSection As String, ByVal szSectionInstall As String)
            Dim szINI As String
            If Me.iResult_IUM <= 0 Then
                szINI = Me.MainINIFile.ReadKey(szSectionInstall, SOME_INIFILE_KEY)
                szINI = Me.strProjectRoot_IUW & "\" & szINI
                Console.WriteLine("INI FIle = {0}", szINI)
                If IO.File.Exists(szINI) Then
                    Dim iniFile As New Ini.IniFile(szINI)
                    Dim szPath, szFileName, szExtension, szPref, szHost, szPort As String
                    Dim szUser, szPass, szMainFolder, szURL, szLocalFile As String
                    Dim szBuild As String
                    szPath = Me.strProjectRoot_IUW & "\" & DISTRIB_ROOT_FOLDER
                    szHost = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPHOST_KEY)
                    szPref = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPREF_KEY)
                    szPort = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPORT_KEY)
                    szMainFolder = iniFile.ReadKey(szSectionInstall, MAINFOLDER_KEY)
                    szUser = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPUSER_KEY)
                    szPass = Me.MainINIFile.ReadKey(HOSTUPDATE_SECTION, HTTPPASS_KEY)
                    szFileName = iniFile.ReadKey(szSection, "filename")
                    szExtension = iniFile.ReadKey(szSection, "extension")
                    If szUser = Nothing Or szUser.Length = 0 Then
                        If StrComp(szPref, "http", vbTextCompare) = 0 Then
                            szURL = szPref & "://" & szHost & ":" & szPort & szMainFolder
                            szLocalFile = szFileName & szExtension
                            Console.WriteLine("URL Path = {0}", szURL)
                            Console.WriteLine("Local Path = {0}", szPath)
                            Console.WriteLine("File: {0}", szLocalFile)
                            If Not DownloadFilesFromIntHttp(szLocalFile, szURL, szPath) = 0 Then
                                Me.iResult_IUM = 6
                            Else
                                szBuild = iniFile.ReadKey(szSection, BUILD_PROGRAM)
                                If StrComp(szExtension, ".exe", 0) = 0 And StrComp(szBuild, BUILD_EXE_JAVA, 0) = 0 Then
                                    Console.WriteLine("Build Program = {0}", BUILD_EXE_JAVA)
                                    Console.WriteLine("Install {0} Package...", szLocalFile)
                                    'MsgBox(Chr(34) & szPath & "\" & szLocalFile & Chr(34) & " /s")
                                    ' Attention!!!
                                    Dim myProcess As Process
                                    myProcess = System.Diagnostics.Process.Start(Chr(34) & szPath & "\" & szLocalFile & Chr(34), " /s")
                                    myProcess.WaitForExit()
                                    myProcess.Close()
                                    Console.WriteLine(" ")
                                    Console.WriteLine("Installation Success!")
                                    Me.iResult_IUM = 0
                                Else
                                    Me.iResult_IUM = 7
                                End If
                            End If
                        Else
                            Me.iResult_IUM = 100
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

        Private Function DownloadFilesFromIntHttp(ByVal strFile As String, ByVal strURL As String, ByVal strPath As String) As Integer
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
            If IO.Directory.Exists(strPath) Then
                intUploadFilesFromInt = 0
            Else
                intUploadFilesFromInt = 1
            End If

            ' If File Exsit Delete it
            If IO.File.Exists(strLocal_Path) Then
                IO.File.Delete(strLocal_Path)
                ' Test
                If IO.File.Exists(strLocal_Path) Then
                    intUploadFilesFromInt = 1
                Else
                    intUploadFilesFromInt = 0
                End If
            End If

            ' **** Download File ****
            Const ITERATION As Integer = 10
            If intUploadFilesFromInt = 0 Then
                For I As Integer = 1 To ITERATION
                    Try
                        Console.WriteLine("Downloading " & strfileURL & ", Attempt #" & Str(I) & " ...")
                        ' Test
                        My.Computer.Network.DownloadFile(strfileURL, strLocal_Path)
                        intUploadFilesFromInt = 0
                        Exit For
                        Exit Try
                    Catch Err01
                        MsgBox("Wrong Destination Path")
                        intUploadFilesFromInt = 1
                        Exit For
                        Exit Try
                    Catch Err02 When I = ITERATION
                        MsgBox("Time out Exception for URI: " & strfileURL & " : " & Err02.Message)
                        intUploadFilesFromInt = 2
                        Exit Try
                    Catch Err02 When I < ITERATION
                        Console.WriteLine("Time out!")
                        intUploadFilesFromInt = 2
                        Continue For
                        Exit Try
                    Catch Err03
                        MsgBox("Security Exception. Anonymous Authentication Denied")
                        intUploadFilesFromInt = 2
                        Exit For
                        Exit Try
                    Catch Err04 When I = ITERATION
                        MsgBox("Web Server Exception. Connection Refused" & " : " & Err04.Message)
                        intUploadFilesFromInt = 2
                        Exit Try
                    Catch Err04 When I < ITERATION
                        Console.WriteLine("Connection refused!")
                        intUploadFilesFromInt = 2
                        Continue For
                        Exit Try
                    End Try
                Next I
            End If


            ' **** /Download File ****

            ' **** Check if Files is Downloaded **** '
            ' Test
            If Not IO.File.Exists(strLocal_Path) And intUploadFilesFromInt = 0 Then
                intUploadFilesFromInt = 1
            End If
            ' **** /Check Path if Exist **** '
            DownloadFilesFromIntHttp = intUploadFilesFromInt
        End Function

#End Region '/МЕТОДЫ РАБОТЫ С INI-ФАЙЛАМИ

#Region "Методы работы с реестром"
        ' *********************************************************************************
        '
        ' Subroutine Win7x86_WriteRegistry
        ' This Subroutine Adds at Startup Registry Section Files Of Setup Initialization
        ' Parameters Loads from INI File
        '
        ' PARAMETERS:	szSection -- the Update File Section
        '
        ' RETURNS:		None
        '
        ' DEPENDS:		Constants.vb, Constants.sysupdate.vb
        '
        ' ****************************************************************************
        Public Sub Win7x86_WriteRegistry(ByVal strSection As String)
            Dim objStreamWriter As StreamWriter
            Dim szRegScriptProgram As String
            Dim szINI As String
            Dim iNum As Integer
            Dim szFullRegProgram As String
            szFullRegProgram = Me.strProjectRoot_IUW & "\" & WRITE_REGISTRYPROGRAM
            Dim strEcho As String
            If IO.File.Exists(szFullRegProgram) Then

                strEcho = "*******************************************************" & vbCrLf & vbCrLf
                strEcho = strEcho & "Ваша Windows 7/Windows Server 2008 R2" & vbCrLf
                strEcho = strEcho & "нуждается в обновлении. Будут установлены" & vbCrLf
                strEcho = strEcho & "обновления: .Net Framework 4.5," & vbCrLf
                strEcho = strEcho & "Windows Managements Framework 5.1," & vbCrLf
                strEcho = strEcho & "Internet Explorer 11, TLS 1.2," & vbCrLf
                strEcho = strEcho & "Some Security Updates, Chocolatey." & vbCrLf
                strEcho = strEcho & "Во время установки программ компьютер несколько" & vbCrLf
                strEcho = strEcho & "раз перезагрузится (до восьми раз). Примерное" & vbCrLf
                strEcho = strEcho & "время установки - один час. Программа сама" & vbCrLf
                strEcho = strEcho & "уведомит Вас об окончании установки. Закройте все" & vbCrLf
                strEcho = strEcho & "работающие программы и перегрузите компьютер" & vbCrLf & vbCrLf
                strEcho = strEcho & "*******************************************************" & vbCrLf
                iNum = 100

                If Me.iResult_IUM > 0 Then
                    Me.iResult_IUM = 0
                End If
                ' If strSection Not Exist (First Run)
                If strSection Is Nothing Or strSection.Length = 0 Then
                    MsgBox(strEcho, vbOKOnly And vbInformation, "Информационное сообщение")
                    ' Test
                    Console.WriteLine("Key = 1" & vbCrLf & "Value = " & Chr(34) & szFullRegProgram & Chr(34) & " " & Chr(34) & UPDATE_NET_SECTION & Chr(34))
                    Registry.SetValue(REGISTRY_KEY_EX, "1", Chr(34) & szFullRegProgram & Chr(34) & " " & Chr(34) & UPDATE_NET_SECTION & Chr(34))
                Else
                    ' Read Update INI file
                    szINI = Me.MainINIFile.ReadKey(strSection, SOME_INIFILE_KEY)
                    szINI = Me.strProjectRoot_IUW & "\" & szINI
                    Console.WriteLine("INI FIle = " & szINI)
                    If IO.File.Exists(szINI) Then
                        '	if File Exist
                        Dim iniFile As New Ini.IniFile(szINI)
                        Dim szKeyTitle As String
                        Dim szNextSection As String
                        Dim szRegProgram As String
                        szKeyTitle = iniFile.ReadKey(strSection, REGISTRY_KEYTITLE)
                        szNextSection = iniFile.ReadKey(strSection, NEXT_SECTIONINSTALL)
                        szRegProgram = iniFile.ReadKey(strSection, REGISTRY_PROGRAM)
                        If szRegProgram Is Nothing Or szRegProgram.Length = 0 Then
                            iResult_IUM = -1
                        Else
                            szRegProgram = Me.strProjectRoot_IUW & "\" & szRegProgram
                            szRegScriptProgram = szRegProgram & ".bat"
                            If IO.File.Exists(szRegScriptProgram) Then
                                IO.File.Delete(szRegScriptProgram)
                            End If
                            objStreamWriter = New StreamWriter(szRegScriptProgram, False, Encoding.GetEncoding(866))
                            ' Test
                            Console.WriteLine("Key1 Value = " & Chr(34) & szRegProgram & Chr(34))
                            'Registry.SetValue(REGISTRY_KEY_EX, "1", Chr(34) & szRegProgram & Chr(34))
                            objStreamWriter.WriteLine(Chr(34) & szRegProgram & Chr(34))
                            If szNextSection Is Nothing Or szNextSection.Length = 0 Then
                                iResult_IUM = -1
                            Else
                                ' Test
                                Console.WriteLine("Key2 Value = " & Chr(34) & szFullRegProgram & Chr(34) & " " & Chr(34) & szNextSection & Chr(34))
                                'Registry.SetValue(REGISTRY_KEY_EX, "2", Chr(34) & szFullRegProgram & Chr(34) & " " & Chr(34) & szNextSection & Chr(34))
                                objStreamWriter.WriteLine(Chr(34) & szFullRegProgram & Chr(34) & " " & Chr(34) & szNextSection & Chr(34))
                            End If
                            ' Attention!!!'
                            objStreamWriter.WriteLine("rem " & "shutdown /r /t 00")
                            objStreamWriter.Close()
                        End If
                    Else
                        'if szINI File Not Exist
                        Me.iResult_IUM = 3
                    End If
                End If
                If iResult_IUM = 0 Then
                    Console.WriteLine("To be Continued. Установка продолжается")
                    If Not szRegScriptProgram Is Nothing Then
                        Registry.SetValue(REGISTRY_KEY_EX, "2", "C:\Windows\system32\cmd.exe /c " & Chr(34) & szRegScriptProgram & Chr(34))
                    End If
                End If
                If iResult_IUM = -1 Then
                    Console.WriteLine("Installation will be canceled. Установка скоро прекратится")
                    If Not szRegScriptProgram Is Nothing Then
                        Registry.SetValue(REGISTRY_KEY_EX, "2", "C:\Windows\system32\cmd.exe /c " & Chr(34) & szRegScriptProgram & Chr(34))
                    End If
                End If
            Else
                Me.iResult_IUM = 3
            End If
        End Sub

#End Region

    End Class

End Namespace


