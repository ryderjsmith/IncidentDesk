[Setup]
AppId={{8A3B5C1E-7F42-4D8E-9A1B-6C3D2E4F5A6B}
AppName=Incident Desk
AppVersion=1.5
AppPublisher=
AppPublisherURL=
AppSupportURL=
AppUpdatesURL=
DefaultDirName={autopf}\IncidentDesk
DefaultGroupName=Incident Desk
AllowNoIcons=yes
OutputDir=..\installer_output
OutputBaseFilename=IncidentDesk-Setup
SetupIconFile=..\img\favicon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
UninstallDisplayIcon={app}\IncidentDesk.exe
UninstallDisplayName=Incident Desk
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "..\dist\IncidentDesk.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Incident Desk"; Filename: "{app}\IncidentDesk.exe"
Name: "{group}\Uninstall Incident Desk"; Filename: "{uninstallexe}"
Name: "{commondesktop}\Incident Desk"; Filename: "{app}\IncidentDesk.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\IncidentDesk.exe"; Description: "Launch Incident Desk"; Flags: nowait postinstall skipifsilent
