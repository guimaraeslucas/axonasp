; AxonASP Server - InnoSetup
; Copyright (C) 2026 G3pix Ltda. All rights reserved.
#define MyAppName "G3pix AxonASP"

;Allows github to change version
#ifndef MyAppVersion
  #define MyAppVersion "2.3.0"
#endif

#define MyAppPublisher "G3pix"
#define MyAppURL "https://g3pix.com.br/axonasp/"
#define MyAppExeName "axonasp-http.exe"

#ifndef MyArchSuffix
  #define MyArchSuffix "amd64"
#endif

#ifndef MyVersionInfoVersion
  #define MyVersionInfoVersion MyAppVersion
#endif

[Setup]
AppId={{8ADC2612-E3A8-4156-85B2-8484565F9C03}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppComments=G3pix is a trademark of G3pix Ltda - Brasil
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
AppCopyright=Copyright (C) 2026 G3pix Ltda.

DefaultDirName=C:\axonasp
UninstallDisplayIcon={app}\{#MyAppExeName}
PrivilegesRequired=admin
ArchitecturesAllowed=x64compatible arm64
ArchitecturesInstallIn64BitMode=x64compatible arm64
DefaultGroupName={#MyAppName}
AllowNoIcons=no

LicenseFile=LICENSE.txt
OutputDir=Output
OutputBaseFilename=axonasp_installer_{#MyAppVersion}_{#MyArchSuffix}
SetupIconFile=resources\icon_server.ico


SolidCompression=yes
Compression=lzma2/ultra
WizardStyle=modern windows11
VersionInfoVersion={#MyVersionInfoVersion}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=AxonASP Server Setup
VersionInfoProductName={#MyAppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "axonasp-http.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "axonasp-admin.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "axonasp-cli.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "axonasp-fastcgi.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "axonasp-mcp.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "axonasp-testsuite.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "axonasp-service.exe"; DestDir: "{app}"; Flags: ignoreversion

Source: "install-service.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "uninstall-service.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "iis-http.cmd"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "LEGAL-DISCLAIMER.md"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "global.asa"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

Source: "resources\icon_service.ico"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist

Source: "config\*"; DestDir: "{app}\config\"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "www\*"; DestDir: "{app}\www\"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "mcp\*"; DestDir: "{app}\mcp\"; Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{group}\{#MyAppName} CLI"; Filename: "{app}\axonasp-cli.exe"; WorkingDir: "{app}"
Name: "{group}\{#MyAppName} TestSuite"; Filename: "{app}\axonasp-testsuite.exe"; WorkingDir: "{app}"
Name: "{group}\{#MyAppName} MCP"; Filename: "{app}\axonasp-mcp.exe"; WorkingDir: "{app}"
Name: "{group}\{#MyAppName} Configuration Manager"; Filename: "{app}\axonasp-admin.exe"; WorkingDir: "{app}"
Name: "{group}\{#MyAppName} Service"; Filename: "{app}\axonasp-service.exe"; WorkingDir: "{app}"
Name: "{group}\{#MyAppName} Install Service"; Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\install-service.ps1"""; IconFilename: "{app}\icon_service.ico"; WorkingDir: "{app}"
Name: "{group}\{#MyAppName} Uninstall Service"; Filename: "powershell.exe"; Parameters: "-ExecutionPolicy Bypass -File ""{app}\uninstall-service.ps1"""; IconFilename: "{app}\icon_service.ico"; WorkingDir: "{app}"
Name: "{group}\{cm:ProgramOnTheWeb,{#MyAppName}}"; Filename: "{#MyAppURL}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon


[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; WorkingDir: "{app}"; Flags: nowait runascurrentuser postinstall skipifsilent
Filename: "powershell"; Parameters: "{app}\install-service.ps1"; Description: "Install and activate windows service"; WorkingDir: "{app}"; Flags: unchecked nowait postinstall runascurrentuser skipifsilent 

[UninstallRun]
Filename: "powershell"; Parameters: "{app}\uninstall-service.ps1"; WorkingDir: "{app}"; RunOnceId: "DelService"; Flags: runascurrentuser skipifdoesntexist

[Registry]
Root: HKLM; Subkey: "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{app}"; Check: NeedsAddPath('{app}'); Flags: uninsdeletevalue

[Code]
function NeedsAddPath(Param: string): boolean;
var
  OrigPath: string;
begin
  if not RegQueryStringValue(HKLM, 'SYSTEM\CurrentControlSet\Control\Session Manager\Environment', 'Path', OrigPath) then begin
    Result := True;
    exit;
  end;
  Result := Pos(';' + Uppercase(Param) + ';', ';' + Uppercase(OrigPath) + ';') = 0;
end;