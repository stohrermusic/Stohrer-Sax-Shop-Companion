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
AppCopyright=Copyright (C) Matt Stohrer
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
ArchitecturesInstallIn64BitMode=x64compatible
; Require Windows 10 or newer (Python 3.11 / PyInstaller drop support below this).
MinVersion=10.0

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

; Custom wording on the welcome and finish pages. Most existing users have
; been running a bare SaxShopCompanion-Windows.exe (downloaded into a
; folder of their choosing — Downloads, Desktop, etc.) since the installer
; only just became the published Windows distribution. Tell them they can
; safely delete the old .exe and that their saved data will carry over.
[Messages]
WelcomeLabel2=This will install Matt Stohrer's "Sax Shop Companion" on to your PC. I guess they let anyone put out software now?%n%nAnyways, PLEASE NOTE if you are upgrading from a previous bare .exe that you dropped in a folder, this version will not remove it, only add itself to your PC. If you want the old one gone, delete it. Your old settings should "just work".
FinishedLabel=Setup has finished installing [name] on your computer. The application may be launched by selecting the installed shortcuts.%n%nIf you previously ran this app as a downloaded SaxShopCompanion-Windows.exe, you can delete that file now — your settings have already been picked up by the installed copy. The Start Menu (and optional desktop) shortcut you just installed is now the canonical way to launch the app.%n%nClick Finish to exit Setup.

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
