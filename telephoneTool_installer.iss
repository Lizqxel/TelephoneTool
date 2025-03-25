; TelephoneTool用インストーラースクリプト
; Inno Setup用スクリプトファイル

#define MyAppName "TelephoneTool"
#define MyAppVersion "1.0.1"
#define MyAppPublisher "Lizqxel"
#define MyAppExeName "TelephoneTool-1.0.1.exe"

[Setup]
; アプリケーション基本情報
AppId={{A1B2C3D4-E5F6-4747-8899-AABBCCDDEEFF}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL="https://example.com/"
AppSupportURL="https://example.com/"
AppUpdatesURL="https://example.com/"
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
; インストーラーがレジストリに作成する内容
DisableProgramGroupPage=yes
; 圧縮設定
Compression=lzma
SolidCompression=yes
; 出力ファイル名設定
OutputDir=installer
OutputBaseFilename={#MyAppName}-{#MyAppVersion}-setup
; Windows 8〜11のユーザーインターフェイスガイドラインに準拠
WizardStyle=modern
; アイコン設定
SetupIconFile=map.png

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; ビルドされたアプリケーション全体をコピー
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "settings.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "map.png"; DestDir: "{app}"; Flags: ignoreversion
Source: "qt.conf"; DestDir: "{app}"; Flags: ignoreversion
Source: "version.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "requirements.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "utils\*"; DestDir: "{app}\utils"; Flags: ignoreversion recursesubdirs
Source: "ui\*"; DestDir: "{app}\ui"; Flags: ignoreversion recursesubdirs
Source: "services\*"; DestDir: "{app}\services"; Flags: ignoreversion recursesubdirs

[Icons]
; スタートメニューとデスクトップにショートカットを作成
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; インストール完了後にアプリケーションを実行するかどうかの選択肢を表示
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent 