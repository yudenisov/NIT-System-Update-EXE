[Setup]
AppName=Win7x86
AppVersion=1.0.0.0.0
AppCopyright=New Internet Technologies Inc.
RestartIfNeededByRun=False
AllowCancelDuringInstall=False
DefaultDirName=C:\ProgramData\NIT\7x86
AllowRootDirectory=True
SolidCompression=True
OutputDir=D:\Download
OutputBaseFilename=Nit.Win7x86
ArchitecturesAllowed=x86
MinVersion=0,6.1
OnlyBelowVersion=0,6.2

[Files]
Source: "Sources_Lib\ie11_install.ini"; DestDir: "{app}"
Source: "Sources_Lib\ie11_update.ini"; DestDir: "{app}"
Source: "Sources_Lib\main.ini"; DestDir: "{app}"
Source: "Sources_Lib\net_framework.ini"; DestDir: "{app}"
Source: "Sources_Lib\other_update.ini"; DestDir: "{app}"
Source: "Sources_Lib\secupdate_01.ini"; DestDir: "{app}"
Source: "Sources_Lib\tls_1.2.ini"; DestDir: "{app}"
Source: "Sources_Lib\winmf_5.1.ini"; DestDir: "{app}"
Source: "NIT.SYSUPDATE.Library\NIT.SYSUPDATE.Library\bin\Release\NIT.SYSUPDATE.Library.dll"; DestDir: "{app}"
Source: "IE11.Install\IE11.Install\bin\Release\IE11.Install.exe"; DestDir: "{app}"
Source: "IE11_Update.Install\IE11_Update.Install\bin\Release\IE11_Update.Install.exe"; DestDir: "{app}"
Source: "NET_Framework.Install\NET_Framework.Install\bin\Release\NET_Framework.Install.exe"; DestDir: "{app}"
Source: "Others.Install\Others.Install\bin\Release\Others.Install.exe"; DestDir: "{app}"
Source: "SecUpdate01.Install\SecUpdate01.Install\bin\Release\SecUpdate01.Install.exe"; DestDir: "{app}"
Source: "Solution1\regprogram.exe\bin\Release\regprogram.exe"; DestDir: "{app}"
Source: "TLS_1.2.Install\TLS_1.2.Install\bin\Release\TLS_1.2.Install.exe"; DestDir: "{app}"
Source: "WinMF_5.1.Install\WinMF_5.1.Install\bin\Release\WinMF_5.1.Install.exe"; DestDir: "{app}"
Source: "Sources_Lib\delete_uninstall.cmd"; DestDir: "{app}"
Source: "Sources_Lib\regprogram.vbs"; DestDir: "{app}"
Source: "Sources_Lib\runregprogram.vbs"; DestDir: "{app}"

[Dirs]
Name: "{app}\Distrib"

[UninstallRun]
Filename: "{app}\delete_uninstall.cmd"; WorkingDir: "{app}"; Flags: shellexec skipifdoesntexist runhidden

[Run]
Filename: "{app}\regprogram.exe"
