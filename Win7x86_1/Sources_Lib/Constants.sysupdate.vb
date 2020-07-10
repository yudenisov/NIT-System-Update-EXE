' *****************************************************************************
'
' Module with Constants for NIT_System_Update Functions
'
' *****************************************************************************

Module Constants_Sysupdate
    ' INI File Keys Constants
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
    Public Const UPDATE_WMF_SECTION As String = "WinMF 5.1"
    Public Const UPDATE_TLS_SECTION As String = "TLS 1.2"
    Public Const UPDATE_SEC_SECTION As String = "SecUpdate 01"
    Public Const UPDATE_OTHER_SECTION As String = "Other Update"
    Public Const DISTRIB_FOLDER_KEY As String = "folder"
    Public Const SOME_INIFILE_KEY As String = "ini_file"
    Public Const REGISTRY_KEYTITLE As String = "keytitle"
    Public Const REGISTRY_PROGRAM As String = "regprogram"
	Public Const NEXT_SECTIONINSTALL As String = "next_sectioninstall"
	Public Const BUILD_PROGRAM As String = "build"
	Public Const NAME_MSI_PACKETS As String = "productname"
	Public Const INITIAL_SECTION As String = "ini_section"

    ' Registry Keys Constants
    Public Const REGISTRY_KEY_EX As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    Public Const REGISTRY_KEY_OR As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    Public Const REGISTRY_TITLE_MAIN As String = "Установка приложений"
    Public Const REGISTRY_TITLE_REBOTING As String = "Cleaning Up and Rebooting"
    ' Main Distrib Downloaded to Folder
    Public Const DISTRIB_ROOT_FOLDER As String = "Distrib"
    ' Program Names
    Public Const PROGRAM_NETFRAMEWORK As String = "NET_Framework.Install.exe"
	'
	' Build Program Variables
	Public Const BUILD_EXE_INSTSHELL As String = "exe_instshell"
	Public Const BUILD_MSU_INSTSHELL As String = "msu_instshell"
	Public Const BUILD_MSI_INSTSHELL As String = "msi_instshell"
	Public Const BUILD_EXE_INNOSETUP As String = "exe_innnosetup"
	Public Const BUILD_EXE_JAVA As String = "exe_java"
	Public Const BUILD_BAT As String = "bat_shell"
	' Names of the Main Functions
	Public Const WRITE_REGISTRYPROGRAM As String = "regprogram.exe"
    Public Const WRITE_REGISTRYSCRIPT As String = "regprogram.vbs"	
End Module
