; TelephoneTool用インストーラースクリプト
; Inno Setup用スクリプトファイル

#define MyAppName "TelephoneTool"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "TelephoneTool"
#define MyAppURL "https://example.com/"
#define MyAppExeName "TelephoneTool.exe"

[Setup]
; アプリケーション基本情報
AppId={{E1FD7AE3-B14D-4F54-A673-F29D90E68C9C}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
; インストーラーがレジストリに作成する内容
DisableProgramGroupPage=yes
; 圧縮設定
Compression=lzma
SolidCompression=yes
; 出力ファイル名設定
OutputDir=installer
OutputBaseFilename=TelephoneTool_Setup
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
Source: "dist\TelephoneTool\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; スタートメニューとデスクトップにショートカットを作成
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; インストール完了後にアプリケーションを実行するかどうかの選択肢を表示
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent 