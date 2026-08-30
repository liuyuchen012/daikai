; AgoraIn 集控服务器（Windows）安装脚本（Inno Setup 6/7）
; 用法：ISCC.exe AgoraIn-Server.iss
; 产物：AgoraIn-Server-Setup-v3.2.0.exe（自包含单文件发布包，无需预装 .NET）
; 依赖发布产物：publish\Server.win-x64（dotnet publish Server/CheckIn.Server.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o publish/Server.win-x64）

#define MyAppName "AgoraIn Server"
#define MyAppNameCn "AgoraIn 集控服务器"
#define MyAppVersion "3.2.0"
#define MyAppPublisher "LiuYuchen"
#define MyAppExeName "CheckIn.Server.exe"
; 沿用 v2.8.0 服务端安装器的 AppId，使新版安装器可直接覆盖升级
#define MyAppId "{{7C2B9E3D-4A6F-4C1E-9D8B-2A5F0E6C1B4A}"

[Setup]
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/liuyuchen012/AgoraIn
AppSupportURL=https://doc.615mc.cn
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=AgoraIn-Server-Setup-v{#MyAppVersion}
OutputDir=..\..\publish\installer
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0
SetupIconFile=icon.ico
LicenseFile=LICENSE.rtf
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppNameCn}
VersionInfoVersion={#MyAppVersion}
VersionInfoProductName={#MyAppNameCn}
VersionInfoProductVersion={#MyAppVersion}
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"

[Files]
; 服务端自包含单文件发布产物（排除 PDB 与运行期数据，保留 config.json 与数据库）
Source: "..\..\publish\Server.win-x64\*"; DestDir: "{app}"; Excludes: "*.pdb,config.json,*.db,*.db-shm,*.db-wal,logs"; Flags: ignoreversion recursesubdirs createallsubdirs

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务"

[Icons]
Name: "{group}\{#MyAppNameCn}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppNameCn}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{group}\卸载 {#MyAppNameCn}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即启动 {#MyAppNameCn}"; Flags: nowait postinstall skipifsilent
