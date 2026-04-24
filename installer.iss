; Inno Setup script for Stohrer Sax Shop Companion (Windows).
; Wraps the PyInstaller-built dist\SaxShopCompanion.exe into an installer
; that places the app under Program Files, adds Start Menu / optional
; desktop shortcuts, and registers an uninstaller.
;
; Build locally (after `python build.py`):
;   iscc /DAppVersion=2.0.1 installer.iss
;
; CI passes /DAppVersion automatically from config.py's APP_VERSION.

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

#define AppName "Stohrer Sax Shop Companion"
#define AppPublisher "Matt Stohrer"
#define AppURL "https://www.stohrermusic.com"
#define AppExeName "SaxShopCompanion.exe"

[Setup]
; Stable product GUID — do NOT change between releases, or Windows will
; treat upgrades as a different product and leave the old install behind.
AppId={{F7A4E3B1-5C2D-4A9F-8B1E-3D6C8E5A2F91}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
DefaultDirName={autopf}\SaxShopCompanion
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=SaxShopCompanion-Windows-Setup-{#AppVersion}
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}
Compression=lzma2/ultra
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "dist\SaxShopCompanion.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent
