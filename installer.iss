; Inno Setup script - Sistem Penugasan Pegawai
; Letakkan file ini di folder root proyek, sejajar dengan main.py.
; Build aplikasi menggunakan PyInstaller sebelum melakukan Compile.

#define MyAppName "Sistem Penugasan Pegawai"
#define MyAppVersion "1.1.1"
#define MyAppPublisher "Briyan Yehezkhiel"
#define MyAppURL "https://github.com/briyanyehezkhiel/SistemPenugasan"
#define MyAppExeName "SistemPenugasan.exe"
#define MyAppSourceDir "dist\SistemPenugasan"

[Setup]
; Pertahankan AppId ini agar installer dikenali sebagai pembaruan aplikasi yang sama.
AppId={{CFC8C9E9-0953-412F-A42E-B1278DA2BCEC}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=SistemPenugasan-Setup-v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Sertakan seluruh hasil build PyInstaller --onedir, termasuk subfolder dan file pendukung.
Source: "{#MyAppSourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
