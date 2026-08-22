;===========================================================
; AgoraIn WPF Client — Inno Setup 安装脚本
;===========================================================

#define AppName       "AgoraIn"
#define AppVersion    "2.8.34"
#define AppPublisher  "AgoraIn"
#define AppURL        "https://github.com/your-repo/check-in-net"
#define AppExeName    "AgoraIn.exe"
#define SourcePath    "Client.win-x64"
#define IconFile      "Client.win-x64\icon.ico"

[Setup]
; 安装程序基本信息
AppId={{8F3A1C62-7D4E-4B90-A1F5-E8D2C9B6A407}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}

; 默认安装目录
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}

; 输出目录与文件名
OutputDir=installer
OutputBaseFilename=AgoraIn-Setup-v{#AppVersion}

; 安装程序图标
SetupIconFile={#IconFile}

; 压缩方式
Compression=lzma2/ultra64
SolidCompression=yes

; 架构支持
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

; 最低 Windows 版本 (Windows 10 1809+)
MinVersion=10.0.17763

; 管理员权限（按需，此处使用 lowest 让用户可选）
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; 界面设置
WizardStyle=modern
DisableProgramGroupPage=yes

; 安装许可协议
LicenseFile=LICENSE.rtf

; 卸载时自动关闭应用
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "chinese"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
; 主程序及全部发布文件（包含 .NET 运行时和所有依赖 DLL）
Source: "{#SourcePath}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; 开始菜单快捷方式
Name: "{group}\{#AppName}";            Filename: "{app}\{#AppExeName}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"

; 桌面快捷方式（可选）
Name: "{autodesktop}\{#AppName}";      Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
; 安装完成后询问是否启动程序
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 清理程序生成的运行时数据（如果存在）
Type: filesandordirs; Name: "{app}\data"

[Code]
// 安装前检查是否已有实例在运行
function InitializeSetup: Boolean;
begin
  // 尝试关闭已运行的实例
  if CheckForMutexes('{#AppName}') then
  begin
    if MsgBox('检测到 {#AppName} 正在运行。' + #13#10 +
              '请关闭程序后重试，或点击"确定"自动关闭。',
              mbConfirmation, MB_OKCANCEL) = IDOK then
    begin
      Result := True;
    end
    else
    begin
      Result := False;
      Exit;
    end;
  end;
  Result := True;
end;
