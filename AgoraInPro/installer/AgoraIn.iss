; AgoraIn 桌面端 Windows 安装脚本（Inno Setup 6/7）
; 用法：ISCC.exe AgoraIn.iss
; 产物：AgoraIn-Setup-v3.2.0.exe（自包含发布包，无需预装 .NET）
; 依赖发布产物：publish\Client.win-x64（dotnet publish AgoraInPro/CheckIn.Client.csproj -c Release -r win-x64 --self-contained true -o publish/Client.win-x64）

#define MyAppName "AgoraIn"
#define MyAppNameCn "AgoraIn 桌面端"
#define MyAppVersion "3.2.5"
#define MyAppPublisher "LiuYuchen"
#define MyAppExeName "AgoraIn.exe"
; 沿用 v2.8.34 安装器的 AppId，使新版安装器可直接覆盖升级
#define MyAppId "{{8F3A1C62-7D4E-4B90-A1F5-E8D2C9B6A407}"

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
OutputBaseFilename=AgoraIn-Setup-v{#MyAppVersion}
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
; 客户端自包含发布产物（排除 PDB 与运行期数据，保留用户已有配置与数据）
Source: "..\..\publish\Client.win-x64\*"; DestDir: "{app}"; Excludes: "*.pdb,config.json,workspace.json,*.db,*.db-shm,*.db-wal,logs"; Flags: ignoreversion recursesubdirs createallsubdirs

[Tasks]
Name: "desktopicon"; Description: "创建桌面快捷方式"; GroupDescription: "附加任务"

[Icons]
Name: "{group}\{#MyAppNameCn}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppNameCn}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon
Name: "{group}\卸载 {#MyAppNameCn}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即启动 {#MyAppNameCn}"; Flags: nowait postinstall skipifsilent
