#define MyAppName "Keyzar Fenix Bot"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Uni Design USA"
#define MyAppExeName "KeyzarFenixBot.exe"

[Setup]
AppId={{2E8943BC-43E8-4BFC-9E93-5D0F59A7404D}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\KeyzarFenixBot
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=installer_output
OutputBaseFilename=KeyzarFenixBot_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
RestartApplications=no
UninstallDisplayName={#MyAppName}
SetupLogging=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "dist\KeyzarFenixBot\KeyzarFenixBot.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\KeyzarFenixBot\_internal\*"; DestDir: "{app}\_internal"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dist\KeyzarFenixBot\bot\*"; DestDir: "{app}\bot"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "dist\KeyzarFenixBot\updater.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "dist\KeyzarFenixBot\.env"; DestDir: "{app}"; Flags: onlyifdoesntexist skipifsourcedoesntexist

[Dirs]
Name: "{app}\data"
Name: "{app}\data\logs"
Name: "{app}\data\pending_reports"
Name: "{app}\data\sent_reports"
Name: "{app}\data\screenshots"
Name: "{app}\playwright_profile"

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; WorkingDir: "{app}"; Flags: nowait postinstall skipifsilent

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;
