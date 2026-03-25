[Setup]
AppName=Incident Desk
AppVersion=1.0
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
