# 班级呼叫系统（ClassCall）

一个基于 **ClassIsland** 的班级呼叫系统，用于教室一体机接收老师的文字广播呼叫，并通过 ClassIsland 的朗读功能把呼叫内容朗读出来。

## 三端架构

```
┌──────────────┐     HTTP      ┌──────────────┐     HTTP(轮询)     ┌───────────────────┐
│  呼出端(控制端) │ ───────────▶ │   服务端      │ ◀───────────────── │  插件端(被控端)    │
│  CallCenter   │  发送呼叫     │  CallServer  │   拉取/确认        │  CallPlugin       │
│(Avalonia 桌面) │              │ (ASP.NET)    │                   │ (ClassIsland 插件)│
└──────────────┘              └──────────────┘                   └───────────────────┘
      老师操作                        中转+存储                       教室一体机
                                                                   弹窗 + ClassIsland朗读
```

| 端 | 项目 | 技术栈 | 运行位置 |
|----|------|--------|----------|
| 呼出端（控制端） | `CallCenter/` | Avalonia 跨平台桌面, net10.0 | 教师电脑（Windows / macOS / Linux 均可） |
| 服务端 | `CallServer/` | ASP.NET Core + EF Core Sqlite, net10.0 | 校内服务器/任意电脑（跨平台） |
| 插件端（被控端/呼叫端） | `CallPlugin/` | ClassIsland 插件, net8.0 + Avalonia | 教室一体机（ClassIsland 内） |

## 功能

- **呼出端**：填写服务端地址连接后，实时查看各教室设备在线状态；选择呼叫类型（紧急呼叫 / 普通通知 / 仅朗读）、目标（全部设备或指定教室），一键发送；底部保留呼叫记录。
- **服务端**：设备注册与心跳（30 秒内未心跳视为离线）；呼叫消息中转（广播或指定设备）；呼叫送达确认（ack），避免被控端重复接收。
- **插件端**：安装到 ClassIsland 后，首次启动自动向服务端注册本机（设备名默认计算机名，建议在设置页填教室名）；每 5 秒轮询一次待处理呼叫；收到呼叫后**通过 ClassIsland 朗读功能朗读呼叫内容**，紧急/通知类型同时弹出置顶提醒窗口。

### 呼叫类型

| 类型 | 说明 |
|------|------|
| `urgent` 紧急呼叫 | 红色置顶弹窗 + 朗读 |
| `notice` 普通通知 | 蓝色置顶弹窗 + 朗读 |
| `speech` 仅朗读 | 不弹窗，仅朗读内容 |

## 构建与发布

### 开发调试

```bash
# 服务端（首次自动建库）
cd CallServer
dotnet run

# 呼出端（Windows / macOS / Linux 均可）
cd CallCenter
dotnet run

# 插件端（生成 .cipx 插件包）
cd CallPlugin
dotnet build -c Release
```

### 生成构建产物（Windows 部署）

```powershell
cd ClassCall
dotnet publish .\CallCenter\CallCenter.csproj -c Release -r win-x64 --self-contained true -o .\publish\CallCenter-win-x64
dotnet publish .\CallServer\CallServer.csproj -c Release -r win-x64 --self-contained true -o .\publish\CallServer-win-x64
dotnet build .\CallPlugin\CallPlugin.csproj -c Release
```

产物统一输出到 `ClassCall/publish/`：

| 产物 | 说明 |
|------|------|
| `publish/CallCenter-win-x64/` | 呼出端（自包含，无需安装 .NET，直接运行 `CallCenter.exe`） |
| `publish/CallServer-win-x64/` | 服务端（自包含，直接运行 `CallServer.exe`） |
| `publish/CallPlugin/CallPlugin.cipx` | ClassIsland 插件包（拖入 ClassIsland 安装） |

### Windows 安装包

`installer/ClassCall.iss` 为 Inno Setup 6 安装脚本（借鉴 AgoraInPro 的图标与许可协议），
安装包包含两个组件：**呼出端**（必选）与**服务端**（可选，仅服务器勾选）。

编译安装包（需安装 Inno Setup 6，或将其 `ISCC.exe` 加入 PATH）：

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" .\installer\ClassCall.iss
```

生成 `publish/installer/ClassCall-Setup-1.0.0.exe`。

## 部署步骤

1. **启动服务端**（任意电脑，默认监听 `http://0.0.0.0:5260`，可用 `appsettings.json` 或 `--urls` 覆盖，首次启动自动创建 `classcall.db`）。局域网内建议使用电脑 IP，如 `http://192.168.1.100:5260`。
2. **呼出端**（教师电脑）：填写服务端地址，点击「连接」，选择目标设备并发送呼叫。
3. **插件端**（教室一体机）：将 `CallPlugin.cipx` 拖入 ClassIsland 的插件设置页安装；然后在 **设置 → 插件设置 → 班级呼叫系统** 中填写服务端地址并保存（会自动注册设备），即可开始接收呼叫。

## 通信协议

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/devices/register` | 插件注册，`{name, room}` → `{uuid, password}`（同名复用凭证） |
| POST | `/api/devices/heartbeat` | 插件心跳，更新在线状态 |
| GET | `/api/devices` | 设备列表（含在线状态），呼出端用 |
| POST | `/api/calls/send` | 发送呼叫，`{targetUuid?, type, title, message, sender}` |
| POST | `/api/calls/pull` | 插件轮询待处理呼叫，`{uuid, password}` → `{calls[]}` |
| POST | `/api/calls/ack` | 确认已送达，`{id, uuid, password}` |

## 朗读实现说明

插件端通过 ClassIsland 主程序的朗读服务接口 `ClassIsland.Core.Abstractions.Services.SpeechService.ISpeechService` 实现：

```csharp
var svc = (ISpeechService)services.GetService(typeof(ISpeechService));
svc.EnqueueSpeechQueue(text);   // 将文字加入 ClassIsland TTS 队列，顺序朗读
```

服务容器经 `ClassIsland.Core.AppBase.Current.Services` 获取（运行时与主程序加载的 ClassIsland.Core 程序集保持一致），文字加入 ClassIsland 的 TTS 队列，由 ClassIsland 内置语音按序朗读，无需额外引入语音库。
