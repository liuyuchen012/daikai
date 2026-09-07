; 班级呼叫系统（ClassCall）Windows 安装脚本
; 使用 Inno Setup 6 编译：ISCC.exe ClassCall.iss
; 产物：呼出端（CallCenter，必选）+ 服务端（CallServer，可选）

#define MyAppName "班级呼叫系统"
#define MyAppNameEn "ClassCall"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "ClassCall"
#define MyAppExeName "CallCenter.exe"
#define MyAppServerExe "CallServer.exe"
#define MyAppId "{{8C9B5B1E-3D2A-4A5B-9C1F-6E7A8D9B0C1D}"

[Setup]
; 安装程序本身
AppId={#MyAppId}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppNameEn}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=ClassCall-Setup-{#MyAppVersion}
OutputDir=..\publish\installer
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0

; 使用 AgoraInPro 的 logo 与许可协议
SetupIconFile=icon.ico
LicenseFile=LICENSE.rtf
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}

; 安装过程可选组件（服务端仅部署在服务器时勾选）
[Components]
Name: "center"; Description: "呼出端（控制端）"; Types: full compact; Flags: fixed
Name: "server"; Description: "服务端（仅需部署在服务器时勾选）"; Types: full

[Types]
Name: "full"; Description: "完整安装（呼出端 + 服务端）"
Name: "compact"; Description: "仅呼出端"

[Files]
; 呼出端（Avalonia 自包含发布产物）
Source: "..\publish\CallCenter-win-x64\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: center
; 服务端（ASP.NET Core 自包含发布产物）
Source: "..\publish\CallServer-win-x64\*"; DestDir: "{app}\server"; Flags: ignoreversion recursesubdirs createallsubdirs; Components: server

[Icons]
Name: "{group}\{#MyAppName}（呼出端）"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"; Components: center
Name: "{group}\{#MyAppName}（服务端）"; Filename: "{app}\server\{#MyAppServerExe}"; IconFilename: "{app}\{#MyAppExeName}"; Components: server
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即启动 {#MyAppName} 呼出端"; Flags: nowait postinstall skipifsilent; Components: center

[UninstallDelete]
Type: filesandordirs; Name: "{app}\classcall.db"
Type: filesandordirs; Name: "{app}\server\classcall.db"
