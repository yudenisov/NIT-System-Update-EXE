[Setup]
AppName=Win7x64_Debug
AppVersion=1.0.0.0.0
AppCopyright=New Internet Technologies Inc.
RestartIfNeededByRun=False
AllowCancelDuringInstall=False
DefaultDirName=C:\ProgramData\NIT\7x64
AllowRootDirectory=True
SolidCompression=True
OutputDir=D:\Download
OutputBaseFilename=Win7x64_Demo

[Files]
Source: "Sources_Lib\ie11_install.ini"; DestDir: "{app}"
Source: "Sources_Lib\ie11_update.ini"; DestDir: "{app}"
Source: "Sources_Lib\main.ini"; DestDir: "{app}"
Source: "Sources_Lib\net_framework.ini"; DestDir: "{app}"
Source: "Sources_Lib\other_update.ini"; DestDir: "{app}"
Source: "Sources_Lib\secupdate_01.ini"; DestDir: "{app}"
Source: "Sources_Lib\tls_1.2.ini"; DestDir: "{app}"
Source: "Sources_Lib\winmf_5.1.ini"; DestDir: "{app}"
Source: "NIT.SYSUPDATE.Library\NIT.SYSUPDATE.Library\bin\Debug\NIT.SYSUPDATE.Library.dll"; DestDir: "{app}"
Source: "IE11.Install\IE11.Install\bin\Debug\IE11.Install.exe"; DestDir: "{app}"
Source: "IE11_Update.Install\IE11_Update.Install\bin\Debug\IE11_Update.Install.exe"; DestDir: "{app}"
Source: "NET_Framework.Install\NET_Framework.Install\bin\Debug\NET_Framework.Install.exe"; DestDir: "{app}"
Source: "Others.Install\Others.Install\bin\Debug\Others.Install.exe"; DestDir: "{app}"
Source: "SecUpdate01.Install\SecUpdate01.Install\bin\Debug\SecUpdate01.Install.exe"; DestDir: "{app}"
Source: "Solution1\regprogram.exe\bin\Debug\regprogram.exe"; DestDir: "{app}"
Source: "TLS_1.2.Install\TLS_1.2.Install\bin\Debug\TLS_1.2.Install.exe"; DestDir: "{app}"
Source: "WinMF_5.1.Install\WinMF_5.1.Install\bin\Debug\WinMF_5.1.Install.exe"; DestDir: "{app}"
Source: "Sources_Lib\delete_uninstall.cmd"; DestDir: "{app}"
Source: "Sources_Lib\regprogram.vbs"; DestDir: "{app}"
Source: "Sources_Lib\runregprogram.vbs"; DestDir: "{app}"

[Dirs]
Name: "{app}\Distrib"

[UninstallRun]
Filename: "{app}\delete_uninstall.cmd"; WorkingDir: "{app}"; Flags: shellexec skipifdoesntexist runhidden

[Run]
Filename: "{app}\runregprogram.vbs"
