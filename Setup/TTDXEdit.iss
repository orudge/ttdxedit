; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define AppVersion "1.20.0000"

[Setup]
AppName=TTDX Editor
AppVerName=TTDX Editor 1.20
AppVersion={#AppVersion}
AppPublisher=Owen Rudge
AppPublisherURL=http://www.transporttycoon.net/
AppSupportURL=http://www.transporttycoon.net/
AppUpdatesURL=http://www.transporttycoon.net/
DefaultDirName={pf}\Owen Rudge\TTDX Editor
DefaultGroupName=Owen Rudge\TTDX Editor
InfoBeforeFile=..\Readme.rtf
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
RestartIfNeededByRun=yes
MinVersion=5.1sp3

[Tasks]
; NOTE: The following entry contains English phrases ("Create a desktop icon" and "Additional icons"). You are free to translate them into another language if required.
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"

[Files]
Source: "..\TTDXedit.exe"; DestDir: "{app}"
Source: "..\Modified.exe"; DestDir: "{app}"
Source: "..\readme.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\changes.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Elevate.exe"; DestDir: "{app}"
Source: "..\Elevate.dll"; DestDir: "{app}"
Source: "..\TTDXHelp.dll"; DestDir: "{app}"

Source: "..\SGMPlugin\TTDXEdit.dll"; DestDir: "{app}\SGMPlugin"; Flags: regserver

; VB runtime files and OCXes
Source: "vbrun\mscomctl.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "vbrun\ExLvwU.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "vbrun\ExTvwU.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "vbrun\ShBrowserCtlsU.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "vbrun\StatBarU.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "vbrun\TrackBarCtlU.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver

; Visual C++ redistributable
Source: "vbrun\vcredist_x86.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[INI]
Filename: "{app}\Owen's Transport Tycoon Station.url"; Section: "InternetShortcut"; Key: "URL"; String: "http://www.transporttycoon.net/"
Filename: "{app}\The Transport Tycoon Forums.url"; Section: "InternetShortcut"; Key: "URL"; String: "http://www.tt-forums.net/"

[Icons]
Name: "{group}\TTDX Editor"; Filename: "{app}\TTDXedit.exe"
Name: "{group}\Owen's Transport Tycoon Station"; Filename: "{app}\Owen's Transport Tycoon Station.url"
Name: "{group}\The Transport Tycoon Forums"; Filename: "{app}\The Transport Tycoon Forums.url"
Name: "{group}\Uninstall TTDX Editor"; Filename: "{uninstallexe}"
Name: "{userdesktop}\TTDX Editor"; Filename: "{app}\TTDXedit.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\TTDXEdit.exe"; Description: "Launch TTDX Editor"; Flags: nowait postinstall skipifsilent
Filename: "{tmp}\vcredist_x86.exe"; Parameters: "/q /promptrestart"; WorkingDir: "{tmp}"; StatusMsg: "Installing Visual C++ runtime..."

[Registry]
Root: HKLM; Subkey: "Software\Owen Rudge"; Flags: uninsdeletekeyifempty
Root: HKLM; Subkey: "Software\Owen Rudge\InstalledSoftware"; Flags: uninsdeletekeyifempty
Root: HKLM; Subkey: "Software\Owen Rudge\InstalledSoftware\TTDX Editor"; Flags: uninsdeletekeyifempty
Root: HKLM; Subkey: "Software\Owen Rudge\InstalledSoftware"; ValueType: string; ValueName: "TTDX Editor"; ValueData: {#AppVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Owen Rudge\InstalledSoftware\TTDX Editor"; ValueType: string; ValueName: "Version"; ValueData: {#AppVersion}; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Owen Rudge\InstalledSoftware\TTDX Editor"; ValueType: string; ValueName: "Path"; ValueData: "{app}"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Owen Rudge\TTDX Editor"; Flags: uninsdeletekeyifempty
Root: HKLM; Subkey: "Software\Owen Rudge\TTDX Editor"; ValueType: string; ValueName: "Path"; ValueData: "{app}"; Flags: uninsdeletevalue
Root: HKLM; Subkey: "Software\Owen Rudge\TTDX Editor"; ValueType: string; ValueName: "Version"; ValueData: {#AppVersion}; Flags: uninsdeletevalue

[UninstallDelete]
Type: files; Name: "{app}\The Transport Tycoon Forums.url"
Type: files; Name: "{app}\Owen's Transport Tycoon Station.url"

