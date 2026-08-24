# AgoraIn 课堂打卡与集控平台

> **版本：v2.8.34**（AgoraInPro 桌面客户端）
>
> **在线文档：https://doc.615mc.cn**
>
> 一套面向课堂场景的"打卡 + 课时管理 + 集控平台"解决方案，包含桌面客户端（大屏模式 / 控制模式）、本地集控服务器和移动端 App。

---

## 目录

- [项目简介](#项目简介)
- [功能特性](#功能特性)
- [界面预览](#界面预览)
- [项目结构](#项目结构)
- [技术栈](#技术栈)
- [快速开始](#快速开始)
- [使用指南](#使用指南)
- [服务器 API 参考](#服务器-api-参考)
- [配置与数据存储](#配置与数据存储)
- [常见问题](#常见问题)
- [版本历史](#版本历史)
- [License](#license)

---

## 项目简介

AgoraIn 桌面客户端（`AgoraInPro`，主程序 `AgoraIn.exe`）在一个窗口内提供两种工作模式：

- **大屏模式**：课堂投影场景下的学生打卡面板。支持多任务标签页、学生按钮网格一键打卡、打卡排名展示（前三名金/银/铜高亮）、在线模式与服务器实时同步、二维码签到任务。
- **控制模式**：教学管理员的控制中心。内置**课时划消与排课**系统（学生课时管理、月历排课、不排课日、复制排课、按课时自动扣减），以及**集控平台**管理（登录远程服务器，查看设备/任务/考勤/用户）。

配套的本地服务器（`Server`）与移动端（`Client.Mobile`）提供设备注册、数据同步、二维码签到、Web 管理面板与多用户权限管理。

---

## 功能特性

### 大屏模式（打卡）

**任务管理**

- 左侧**任务树**：文件夹 / 任务两级结构，支持**打开、属性、重命名、删除**（右键菜单），选中项浅蓝高亮
- 右侧**多标签栏**：每个任务一个标签页，支持**新建（+）、切换、关闭**；双击标签可重命名
- 工作区状态持久化：标签页列表与活动标签自动保存到 `workspace.json`，下次启动自动恢复

**打卡面板**

- 学生**按钮网格**：行列数可配置（默认 6×6），每个学生一个圆角按钮，点击打卡
- 已打卡学生按钮变蓝（`#4285f4`），未打卡为浅灰；鼠标悬停半透明反馈
- **右键取消打卡**（已打卡学生），支持批量操作（任务树右键"清空打卡记录"）
- 网格自动按 `ButtonRows × ButtonCols` 布局，超出滚动

**打卡排名**

- 左侧 320px 排名卡实时展示**最早打卡**的学生，按打卡时间排序
- 第 1/2/3 名分别**金 / 银 / 铜**高亮，展示打卡时间

**在线模式与同步**

- 连接集控服务器后实时同步打卡数据，顶栏状态行显示服务器在线/离线指示灯（绿 `#34a853` / 红 `#ea4335`）
- 服务端可推送配置变更（配置版本追踪，自动应用）；集控平台发现新版本时自动弹窗提示管理员
- 支持从服务器加载数据、检查服务器状态

**签到任务（二维码签到）**

- 一键**创建签到**：生成带短码的签到链接与二维码，支持设置**签到密码、教室、科目**
- 学生扫码后自动打卡，结果回写服务器，短码唯一索引防冲突

**菜单栏**

- 顶栏蓝色标题栏左侧：**模式下拉框**（大屏模式 / 控制模式）
- **文件**：导出打卡数据、导入打卡数据、清空打卡记录、新建任务
- **远程**：远程服务器设置、创建签到、检查服务器状态、从服务器加载数据
- **设置**：任务设置（当前标签页）、管理员设置（全局）
- **帮助**：Github、检查版本列表、关于

**窗口体验**

- 无边框圆角窗口（DWM API），支持拖拽、最小化 / 最大化 / 关闭
- 白色半透明模式下拉框、现代化圆角右键菜单、SVG 圆形应用图标

### 控制模式（控制中心）

**统计条**

- 顶栏统计条：**总设备数量 / 总任务数量 / 在线设备数量**（蓝色大数值）

**左侧导航**

- **划课 / 设备列表 / 任务中心 / 集控平台列表**，选中项浅蓝高亮

**划课（课时划消与排课）**

- 学生列表管理：添加 / 删除学生，支持设置初始课时数；剩余课时以蓝色徽章展示
- 课时划消：输入课时数，**划消课时**（扣减）或**增加课时**（赠送），可填备注
- 课时记录流水：红扣（-X）绿增（+X），按学生筛选查看
- 月历排课：42 格月历（6 周 × 7 天），今日格带 **"今" 徽章**，选中日期蓝框高亮
- 排课规则：**上课 / 下课时间必填**（HH:mm，下课不得等于上课，支持跨天如 23:00–01:00），重复排课自动拦截
- 不排课日：周末默认提示，可将任意日期**设为不排课日 / 恢复**（自动清空当日排课，粉底"休"标记）
- **复制排课**：把某日排课复制到多个日期，可跳过不排课日
- **自动扣减课时**：开启后每 30 秒检查一次，课程结束后按 `实际时长 × 每小时课时消耗` 自动扣减学生课时（SlotKey 幂等去重，多实例安全）
- 设置项：**每小时课时消耗**（支持小数，如 0.5 / 1 / 1.5，范围 0–24）、**自动划消开关**

**集控平台列表**

- 打开远程打卡服务器控制台（`ServerControlWindow`）：登录、仪表盘、设备、任务、考勤、用户管理
- 查看服务器状态、版本信息，管理员密码通过 SHA256 哈希校验

### 集控服务器（Server）

- ASP.NET + EF Core + **SQLite** 存储，默认监听端口 **5250**（`config.json` 可改）
- 设备注册与管理、打卡数据同步（`/api/sync_data`、`/api/load_data`）、考勤记录
- **Web 管理面板**：多用户登录、加盐 SHA256 密码哈希、权限管理（用户增删改、改密码、设备分配）
- **签到任务**：创建签到、短码生成、签到结果回写、按设备关联
- **移动端 API**：JWT 令牌认证、二维码签到、管理员仪表盘
- 安全特性：未配置 `ServerPassword` 时自动生成随机密码并回写 `config.json`；`DebugMode` 调试开关；旧版明文密钥自动迁移为加密存储

### 移动端（Client.Mobile，MAUI）

**学生端**

- 登录、任务列表 / 详情、**扫码签到**（`StudentScanPage`）、二维码生成、历史记录、考勤详情

**管理员端**

- 登录、仪表盘、任务管理、用户管理、设置、学生网格、创建签到

---

## 界面预览

大屏模式（打卡面板） 
 ![大屏模式主窗口](./大屏模式主窗口.svg) 
控制模式（课时划消与排课）
![控制模式主窗口](./控制模式主窗口.svg) 

---

## 项目结构

```
check-in-net/
├── AgoraInPro/              # 桌面客户端（WPF，主程序 AgoraIn.exe）
│   ├── MainWindow.xaml(.cs)         # 主窗口：任务树、标签栏、打卡面板、模式切换
│   ├── ControlCenterView.cs         # 控制中心视图（统计条、导航、划课页）
│   ├── ServerControlWindow.xaml(.cs)# 集控平台控制窗口
│   ├── ClassHoursPanelControl.xaml  # 课时划消与排课页（学生列表 + 日历 + 排课面板）
│   ├── ScheduleTimeDialog.cs        # 排课时间对话框（上课/下课必填校验）
│   ├── CopyScheduleDialog.cs        # 复制排课对话框
│   ├── ModernDialog.cs              # 现代风格对话框（提示/确认/更新/关于）
│   ├── App.xaml(.cs)                # 应用资源与启动逻辑
│   ├── Models/                      # AppConfig、ClassHourModels、RemoteModels 等
│   ├── ViewModels/                  # MainViewModel、TaskTabViewModel、ClassHoursViewModel 等
│   ├── Services/                    # ServerService、RemoteControlService、ClassHourStore 等
│   └── CheckIn.Client.csproj        # net10.0-windows，AssemblyName=AgoraIn
├── Server/                  # 集控服务器（ASP.NET + EF Core + SQLite）
│   ├── Program.cs                   # 服务器入口（Web 面板 + 移动端 API + 签到服务）
│   ├── CheckIn.Server.csproj
│   └── config.json                  # Port（默认 5250）、ServerName、ServerPassword、DebugMode
├── Client.Mobile/           # 移动端（.NET MAUI）
│   ├── Pages/                       # 登录/签到/二维码/历史/管理/设置等 17 个页面
│   └── AgoraIn.Client.Mobile.csproj
├── 大屏模式主窗口.svg        # 界面示意图（README 预览用）
├── 控制模式主窗口.svg
└── CheckIn.slnx             # 解决方案文件
```

---

## 技术栈

| 组件 | 技术 |
| --- | --- |
| 桌面客户端 | .NET 10 / WPF（`net10.0-windows`），AssemblyName `AgoraIn` |
| 服务器 | ASP.NET Core + Entity Framework Core + SQLite |
| 移动端 | .NET MAUI（Android / iOS / Windows） |
| 二维码 | QRCoder（签到链接 / 二维码生成） |
| 图标 / 界面 | SVG 生成应用图标与界面示意图 |
| 窗口效果 | DWM API 窗口圆角、半透明阴影、DropShadowEffect |
| 数据文件 | System.Text.Json 序列化（config.json / workspace.json / classhours.json） |

---

## 快速开始

### 环境要求

- Windows 10/11
- [.NET 10 SDK](https://dotnet.microsoft.com/)（构建需要）
- 移动端构建需额外安装 .NET MAUI 工作负载：`dotnet workload install maui`

### 构建

```powershell
# 桌面客户端（输出 AgoraIn.exe）
dotnet build AgoraInPro\CheckIn.Client.csproj

# 集控服务器
dotnet build Server\CheckIn.Server.csproj

# 移动端
dotnet build Client.Mobile\AgoraIn.Client.Mobile.csproj
```

### 运行

```powershell
# 启动桌面客户端（默认大屏模式）
.\AgoraInPro\bin\Debug\net10.0-windows\AgoraIn.exe

# 启动集控服务器（默认端口 5250，首次运行自动生成 ServerPassword 并写入 config.json）
dotnet run --project Server\CheckIn.Server.csproj
```

> 提示：构建前如输出文件被占用，请先结束旧进程：
> `Stop-Process -Name AgoraIn -Force`

### 自测模式

桌面客户端内置自测命令，退出码 0 表示全部通过：

```powershell
& .\AgoraInPro\bin\Debug\net10.0-windows\AgoraIn.exe --selftest
```

---

## 使用指南

### 模式切换

点击主窗口标题栏右侧的**模式下拉框**切换：

- **大屏模式**：进入打卡面板（默认）
- **控制模式**：进入控制中心（划课 / 集控平台），支持最小化 / 最大化 / 关闭

### 大屏打卡

1. 左侧任务树选择任务（或点「+」新建标签页，双击标签重命名）
2. 点击学生按钮即可打卡，已打卡学生按钮变蓝
3. 左侧排名区实时显示最早打卡的学生，前三名金/银/铜高亮
4. 右键已打卡学生可**取消打卡**；任务树右键可**清空打卡记录**
5. 在线模式下打卡数据自动同步到服务器（顶栏状态行显示在线/离线）
6. 需要学生扫码时：**远程 → 创建签到**，生成二维码 / 短码链接，可设签到密码、教室、科目

### 课时划消

1. 进入**控制模式 → 划课**
2. 左侧学生列表选中学生（支持添加 / 删除，添加时可设初始课时）
3. 输入课时数，点击**划消课时**（扣减）或**增加课时**（赠送），可填备注
4. 课时记录写入右侧流水（红扣绿增），学生剩余课时以蓝色徽章展示

### 排课

1. 月历中选择日期（选中日期蓝框高亮，今日带 **"今" 徽章**）
2. 从「选择学生」列表点击学生添加排课，**必须填写上课 / 下课时间**（HH:mm，下课不等于上课，支持跨天如 23:00–01:00）才能提交
3. 已排课学生可修改时间或移除；重复排课会被拦截
4. 工具按钮：
   - **设为不排课日 / 恢复**：自动清空当日排课，周末默认提示（粉底"休"标记）
   - **清空排课**：清除选中日期全部排课
   - **复制排课…**：把某日排课复制到多个日期，可跳过不排课日
5. 设置项中开启「自动扣减课时」后，系统每 30 秒检查一次，当日排课结束后自动按 `实际时长 × 每小时课时消耗` 扣减学生课时（幂等去重，不会重复扣）

### 集控平台

1. **控制模式 → 集控平台列表 → 打开控制台**
2. 输入服务器地址（如 `http://192.168.1.100:5250`）并登录（JWT 认证）
3. 查看仪表盘、设备列表、任务中心、考勤记录，管理用户与设备分配

### 服务器部署

1. 将 `Server` 发布产物复制到目标机器，编辑 `config.json` 设置 `Port`、`ServerName`、`ServerPassword`
2. 运行服务器程序，首次启动自动创建 SQLite 数据库与表结构
3. 局域网内客户端「远程 → 远程服务器设置」填入 IP、端口与密码即可连接

---

## 服务器 API 参考

所有接口返回 JSON，非登录接口需携带服务器密码（`ServerPassword`）校验，管理类接口需 JWT 令牌。

| 类别 | 接口 | 说明 |
| --- | --- | --- |
| 状态 | `GET /api/status` | 服务器状态与设备/任务统计 |
| 状态 | `GET /api/version` | 服务器版本号 |
| 状态 | `GET /api/server_update` | 集控平台新版本检查（后台轮询 GitHub） |
| 设备 | `POST /api/register` | 设备注册 |
| 设备 | `GET /api/machines/{uuid}/tasks` | 获取设备关联任务 |
| 设备 | `POST /api/delete_machine` | 删除设备 |
| 数据 | `POST /api/sync_data` | 打卡数据同步 |
| 数据 | `POST /api/load_data` | 从服务器加载数据 |
| 配置 | `POST /api/get_config` | 获取服务器配置 |
| 配置 | `POST /api/update_config` | 更新服务器配置（版本号 +1，触发客户端推送） |
| 配置 | `POST /api/update_machine_config` | 更新设备配置 |
| 配置 | `POST /api/config_applied` | 客户端确认配置已应用 |
| 打卡 | `POST /api/web_punch` | Web 端打卡 |
| 打卡 | `POST /api/web_cancel_punch` | Web 端取消打卡 |
| 打卡 | `POST /api/clear_attendance` | 清空考勤记录 |
| 签到 | `POST /api/create_signin` | 创建签到任务（短码 + 密码 + 教室/科目） |
| 签到 | `POST /api/signin_result` | 签到结果回写 |
| 用户 | `GET/POST /api/users` | 用户列表 / 创建用户 |
| 用户 | `POST /api/users/change-password` | 修改密码 |
| 认证 | `POST /api/auth/login` | 管理员登录（JWT） |
| 认证 | `POST /api/auth/verify` | 验证令牌 |
| 二维码 | `POST /api/qrcode/generate` | 生成签到二维码 |
| 二维码 | `POST /api/qrcode/checkin` | 扫码签到 |
| 移动端 | `GET /api/mobile/dashboard` | 管理员仪表盘 |
| 移动端 | `GET /api/mobile/attendance` | 考勤记录 |
| 移动端 | `GET /api/mobile/tasks`、`POST /api/mobile/tasks/{id}/close` | 任务列表 / 关闭任务 |
| 移动端 | `GET /api/mobile/tasks/{id}/qrcode` | 任务二维码 |
| 移动端 | `GET /api/mobile/devices`、`POST /api/mobile/devices/{uuid}/tasks` | 设备列表 / 设备分配任务 |
| 移动端 | `GET /api/mobile/students/history` | 学生打卡历史 |
| 移动端 | `GET/POST /api/mobile/assignments` | 排课（设备分配）管理 |
| 移动端 | `GET /api/mobile/teachers` | 教师列表 |
| 调试 | `GET /api/debug/status`、`/api/debug/token`、`POST /api/debug/login` | 调试模式（`DebugMode: true` 时生效） |

---

## 配置与数据存储

### 客户端 `config.json`（全局配置）

| 字段 | 默认值 | 说明 |
| --- | --- | --- |
| `School` | `""` | 学校名称 |
| `Nj` | `""` | 年级（年份） |
| `ClassId` | `""` | 班级编号 |
| `Km` | `""` | 课程名称（如"数学"） |
| `ButtonRows` | `6` | 学生按钮网格行数 |
| `ButtonCols` | `6` | 学生按钮网格列数 |
| `OnlineMode` | `true` | 是否启用在线模式（连接服务器同步） |
| `ServerIp` | `""` | 远程服务器 IP |
| `ServerPort` | `5250` | 远程服务器端口 |
| `ServerPassword` | `""` | 远程服务器连接密码 |
| `AdminPasswordHash` | `""` | 管理员密码 SHA256 哈希（空表示无密码） |

### 任务配置（各任务数据目录 `config.json`）

| 字段 | 说明 |
| --- | --- |
| `Name` | 任务名称（如"三（1）班"） |
| `Km` | 课程名称 |
| `ButtonRows` / `ButtonCols` | 按钮行列数 |
| `OnlineMode` | 是否在线模式 |
| `IsSignInTask` | 是否为签到任务 |
| `SignInTaskId` | 签到任务服务器 TaskId（如 `signin_abc123`） |

### 课时数据 `data/classhours.json`

数据版本 **v3**，结构：

| 字段 | 说明 |
| --- | --- |
| `Version` | 数据版本号（v3：排课细分时间 + 自动划消设置） |
| `Students` | 学生列表（姓名、总课时、已划课时、备注、创建时间） |
| `Records` | 课时记录流水（日期、课时数正负、备注、SlotKey 去重键） |
| `Schedule` | 排课数据：日期 → 排课条目（学生 + 上课/下课时间，支持跨天） |
| `OffDays` | 不排课日集合 |
| `HoursPerHour` | 每小时上课消耗课时（支持小数，默认 1） |
| `AutoDeduct` | 是否自动划消课时 |

### 其他文件

| 文件 | 位置 | 内容 |
| --- | --- | --- |
| `workspace.json` | 客户端目录 | 工作区状态：打开的标签页列表、活动标签页（自动恢复） |
| 打卡数据 / 学生名单 | 任务数据目录 | 打卡记录与学生名单 |
| `config.json` | 服务器目录 | `Port`（默认 5250）、`ServerName`、`ServerPassword`、`DebugMode` |
| SQLite 数据库 | 服务器目录 | 设备、任务、考勤、用户、签到任务、设备分配等持久化数据 |

---

## 常见问题

**Q：构建报"文件被占用 / 无法复制 AgoraIn.exe"？**
A：先结束旧进程再构建：`Stop-Process -Name AgoraIn -Force`

**Q：服务器提示"未配置 ServerPassword，已自动生成"？**
A：这是安全默认行为，请打开服务器目录 `config.json` 查看生成的密码，并在客户端「远程服务器设置」中填入。

**Q：排课无法提交？**
A：检查上课/下课时间是否已填写（HH:mm），且下课时间不得等于上课时间；重复排课也会被拦截。

**Q：开启自动扣减后课时未变化？**
A：确认排课的起止时间均有效，且课程已到下课时间；系统每 30 秒检查一次，同一时段不会重复扣减（SlotKey 去重）。

**Q：客户端连不上服务器？**
A：检查服务器端口（默认 5250）是否开放、IP 是否正确，服务器 `config.json` 中的 `ServerPassword` 是否与客户端一致。

---

## 版本历史

### v2.8.34（当前）

**新增**
- 控制模式（控制中心）：统计条、左侧导航、划课页、集控平台列表
- 课时划消系统：学生管理、课时划消/增加 + 备注、课时记录流水（红扣绿增）
- 月历排课：上课/下课时间必填、今日"今"徽章、选中蓝框、跨天课程、重复排课拦截
- 不排课日（自动清空排课、周末默认提示）、复制排课（可跳过不排课日）
- 自动扣减课时：按实际时长 × 每小时课时消耗，SlotKey 幂等去重
- 大屏模式：任务树、多标签页、签到任务（二维码 + 短码 + 密码）、在线同步、右键取消打卡

**UI/体验**
- 模式切换下拉框（白色半透明样式）
- 现代化圆角：输入框、选项卡、右键菜单、DWM 窗口圆角
- SVG 圆形应用图标

**修复**
- 日历选中蓝框跟随日期
- 大屏模式取消打卡失效
- 控制模式显示错误模式名
- 白字继承链问题（下拉框/菜单文字颜色）

---

## License

[GPL V3](./LICENSE)
