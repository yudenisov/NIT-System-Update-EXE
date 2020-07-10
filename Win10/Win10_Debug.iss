[Setup]
AppName=Win10_Debug
AppVersion=1.0.0.0.0
AppCopyright=New Internet Technologies Inc.
RestartIfNeededByRun=False
AllowCancelDuringInstall=False
DefaultDirName=C:\ProgramData\NIT\10
AllowRootDirectory=True
SolidCompression=True
OutputDir=D:\Download
OutputBaseFilename=Win10_Demo

[Files]
Source: "Sources_Lib\main.ini"; DestDir: "{app}"
Source: "Sources_Lib\other_update.ini"; DestDir: "{app}"
Source: "Sources_Lib\secupdate_01.ini"; DestDir: "{app}"
Source: "NIT.SYSUPDATE.Library\NIT.SYSUPDATE.Library\bin\Debug\NIT.SYSUPDATE.Library.dll"; DestDir: "{app}"
Source: "Others.Install\Others.Install\bin\Debug\Others.Install.exe"; DestDir: "{app}"
Source: "SecUpdate01.Install\SecUpdate01.Install\bin\Debug\SecUpdate01.Install.exe"; DestDir: "{app}"
Source: "Solution1\regprogram.exe\bin\Debug\regprogram.exe"; DestDir: "{app}"
Source: "Sources_Lib\delete_uninstall.cmd"; DestDir: "{app}"
Source: "Sources_Lib\regprogram.vbs"; DestDir: "{app}"
Source: "Sources_Lib\runregprogram.vbs"; DestDir: "{app}"

[Dirs]
Name: "{app}\Distrib"

[UninstallRun]
Filename: "{app}\delete_uninstall.cmd"; WorkingDir: "{app}"; Flags: shellexec skipifdoesntexist runhidden

[Run]
Filename: "{app}\runregprogram.vbs"
