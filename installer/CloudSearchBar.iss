[Setup]
AppName=Cloud Search Bar
AppVersion=1.1.1
AppPublisher=Chris Duffy
AppPublisherURL=
AppSupportURL=
DefaultDirName={autopf}\Cloud Search Bar
DefaultGroupName=Cloud Search Bar
; Runs per-user (no admin) by default, but offers a dialog to install system-wide.
; Per-user installs to AppData\Local, system-wide installs to Program Files.
; NOTE: If installed system-wide, editing cloudsearch_bar.ini requires admin rights.
; The Settings UI (coming soon) will handle this gracefully.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\installer_output
OutputBaseFilename=CloudSearchBar_Setup
SetupIconFile=..\icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
; Prevent multiple instances of the installer running at once
AppMutex=CloudSearchBarInstaller

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
; Main executable
Source: "..\CloudSearchBar.exe"; DestDir: "{app}"; Flags: ignoreversion
; Tray icon
Source: "..\icon.png"; DestDir: "{app}"; Flags: ignoreversion
; Config file — only copy if one doesn't already exist (preserves user settings on reinstall)
Source: "..\cloudsearch_bar.ini"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist

[Icons]
; Start Menu shortcuts
Name: "{group}\Cloud Search Bar"; Filename: "{app}\CloudSearchBar.exe"
Name: "{group}\Uninstall Cloud Search Bar"; Filename: "{uninstallexe}"
; Optional desktop shortcut
Name: "{autodesktop}\Cloud Search Bar"; Filename: "{app}\CloudSearchBar.exe"; Tasks: desktopicon

[Run]
; Launch the app after install so it writes the startup registry entry
Filename: "{app}\CloudSearchBar.exe"; Description: "Launch Cloud Search Bar now"; Flags: nowait postinstall skipifsilent

[Registry]
; Remove the startup registry entry on uninstall
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: none; ValueName: "CloudSearchBar"; Flags: deletevalue uninsdeletevalue
