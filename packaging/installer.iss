#define MyAppName "演習投影 統合管理システム"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Mitsuie"
#define MyAppExeName "Random-Presenter-Selector.exe"
#define MyAppDirName "Random-Presenter-Selector"

[Setup]
; アプリケーション一意ID (プロジェクト固有のGUID)
AppId={{AE80DB53-801F-4B8B-A19B-1377BDB2F310}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}

; 管理者権限不要でインストール可能なユーザーローカル領域に配置
DefaultDirName={localappdata}\Programs\{#MyAppDirName}
DefaultGroupName={#MyAppName}
OutputDir=..\dist_installer
OutputBaseFilename={#MyAppDirName}_Setup_v{#MyAppVersion}
SetupIconFile=app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2/ultra64
SolidCompression=yes

; 一般ユーザー権限でインストール可能にする（UAC昇格ダイアログを出さない）
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes

; 【最重要】自動更新時の旧バージョンプロセス自動検知・終了設定
CloseApplications=yes
CloseApplicationsFilter=*.exe
RestartApplications=no

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "..\dist\{#MyAppDirName}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; スタートメニュー
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
; デスクトップショートカット
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; インストール完了後にアプリを起動するチェックボックス
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
