; Windows installer for DocGen (Inno Setup 6).
; Built by .github/workflows/build-windows.yml:
;   ISCC.exe /DAppVersion=0.1.0 packaging\installer.iss
; Installs per user by default (no administrator rights needed, as on most office PCs);
; an administrator can choose "install for all users" in the first dialog.

#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

[Setup]
AppId={{0AEA444E-46F3-41CA-B4A1-8FC1389210C4}
AppName=DocGen
AppVerName=DocGen {#AppVersion} – Tạo hồ sơ Word từ Excel
AppVersion={#AppVersion}
AppPublisher=DocGen
DefaultDirName={autopf}\DocGen
DefaultGroupName=DocGen
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\dist-installer
OutputBaseFilename=DocGen-Setup-{#AppVersion}-x64
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\DocGen.exe
UninstallDisplayName=DocGen – Tạo hồ sơ Word từ Excel
CloseApplications=yes
#ifexist "docgen.ico"
SetupIconFile=docgen.ico
#endif

[Languages]
#ifexist "Vietnamese.isl"
Name: "vi"; MessagesFile: "Vietnamese.isl"
#endif
Name: "en"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Tạo biểu tượng DocGen trên màn hình (Desktop)"; GroupDescription: "Biểu tượng:"

[Files]
Source: "..\dist\DocGen\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\DocGen – Tạo hồ sơ Word từ Excel"; Filename: "{app}\DocGen.exe"
Name: "{group}\Gỡ cài đặt DocGen"; Filename: "{uninstallexe}"
Name: "{autodesktop}\DocGen"; Filename: "{app}\DocGen.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\DocGen.exe"; Description: "Mở DocGen ngay"; Flags: nowait postinstall skipifsilent
