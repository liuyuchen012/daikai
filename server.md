# SignWave 集控平台 - 服务器 API 文档

## 目录

- [1. 概述](#1-概述)
- [2. 配置说明](#2-配置说明)
- [3. 认证机制](#3-认证机制)
- [4. 数据模型](#4-数据模型)
- [5. API 接口一览](#5-api-接口一览)
- [6. 客户端 API 详解](#6-客户端-api-详解)
- [7. Web 面板 API 详解](#7-web-面板-api-详解)
- [8. 认证页面接口](#8-认证页面接口)
- [9. Web 管理面板页面](#9-web-管理面板页面)
- [10. 调用示例](#10-调用示例)

---

## 1. 概述

SignWave 集控平台是一个轻量级的 **学生打卡管理系统** 服务器，基于 ASP.NET Core Minimal API 构建。提供以下能力：

- **客户端管理**：设备注册、RSA 签名验证、打卡数据同步
- **Web 管理面板**：可视化设备管理、打卡排名查看、远程打卡操作
- **Session 认证**：基于 HMAC-SHA256 签名 Cookie 的安全会话管理

### 技术栈

| 组件 | 说明 |
|------|------|
| 框架 | ASP.NET Core Minimal API (.NET 10) |
| 数据库 | SQLite (EF Core) |
| 验证 | RSA PKCS1 SHA256 + 服务器密码 |
| Web 认证 | HMAC-SHA256 签名的 Cookie Session |
| 端口 | 默认 5250（可配置） |

### 启动方式

```bash
# 开发环境
cd Server
dotnet run

# 生产环境（Linux）
cd Server
dotnet publish -c Release -o ./publish
cd publish
nohup dotnet CheckIn.Server.dll > server.log 2>&1 &
```

---

## 2. 配置说明

服务器启动时读取 `Server/config.json` 配置文件：

```json
{
  "Port": 5250,
  "AdminUsername": "admin",
  "AdminPassword": "admin",
  "ServerName": "SignWave 集控平台",
  "ServerPassword": "admin123"
}
```

| 配置项 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| `Port` | int | `5250` | 服务器监听端口 |
| `AdminUsername` | string | `admin` | Web 管理面板登录用户名 |
| `AdminPassword` | string | `admin` | Web 管理面板登录密码 |
| `ServerName` | string | `SignWave 集控平台` | 显示在 Web 页面标题栏的名称 |
| `ServerPassword` | string | `admin123` | 客户端 API 调用的共享密钥 |

---

## 3. 认证机制

系统使用 **两套并行的认证体系**：

### 3.1 客户端 API 认证

所有客户端 API（`/api/*`）采用 **共享密钥 + RSA 签名** 双重验证：

1. **第一层**：请求体中的 `password` 字段必须与 `ServerPassword` 配置一致
2. **第二层**：对于写操作（数据同步、配置更新），需额外验证 RSA 签名
3. **第三层**：对于读操作（加载数据、获取配置），需验证 challenge-签名

### 3.2 Web 面板认证

Web 管理页面（非 `/api/*` 路径）采用 **HMAC-SHA256 签名 Session**：

1. 用户访问任何页面前必须通过 `/login` 页面认证
2. 登录成功后，服务器生成格式为 `username:timestamp:signature` 的令牌
3. 令牌通过名为 `sw_session` 的 HttpOnly Cookie 下发给浏览器
4. 有效期为 **7 天**
5. 签名使用每个进程启动时随机生成的 `sessionSecret`（进程重启后旧 Cookie 自动失效）

### 3.3 路由保护规则

| 路径前缀 | 是否需要 Web 认证 |
|----------|-------------------|
| `/api/*` | 不需要（密码+签名验证） |
| `/login` | 不需要 |
| `/logout` | 不需要 |
| `/static` | 不需要 |
| 其他所有路径 | **需要登录** |

---

## 4. 数据模型

### 4.1 数据库表结构

#### MachineEntity（设备表）

| 字段 | 类型 | 说明 |
|------|------|------|
| `Uuid` | string(64) | 主键，设备唯一标识符（GUID） |
| `Name` | string | 设备名称（如"三年（1）班"） |
| `PublicKey` | string | 客户端生成的 RSA 公钥（PEM 格式） |
| `LastSeen` | string? | 最后在线时间（ISO 8601 格式） |
| `Config` | string | 客户端任务配置 JSON，结构见 `ClientConfig` |

#### AttendanceEntity（打卡记录表）

| 字段 | 类型 | 说明 |
|------|------|------|
| `Id` | int | 主键，自增 |
| `MachineUuid` | string(64) | 所属设备 UUID（外键） |
| `TaskId` | string(64) | 任务 ID，区分同一设备的不同打卡任务 |
| `Data` | string | 打卡数据 JSON，结构为 `Dictionary<string, StudentAttendance>` |
| `UpdatedAt` | string | 更新时间（ISO 8601 格式） |

> **索引**：`MachineUuid` 单列索引 + `(MachineUuid, TaskId)` 复合索引，优化查询性能。

### 4.2 业务数据模型

#### StudentAttendance（学生打卡数据）

```json
{
  "Name": "张三",
  "Count": 3,
  "FirstTime": "2025-01-15 08:30:00",
  "History": [
    "2025-01-15 08:30:00",
    "2025-01-16 08:25:00",
    "2025-01-17 08:35:00"
  ]
}
```

| 字段 | 类型 | 说明 |
|------|------|------|
| `Name` | string | 学生姓名 |
| `Count` | int | 累计打卡次数 |
| `FirstTime` | string? | 首次打卡时间，`null` 表示从未打卡 |
| `History` | List\<string\> | 历次打卡时间列表 |

#### ClientConfig（客户端配置）

```json
{
  "School": "XX实验中学",
  "Nj": "三年级",
  "ClassId": "1班",
  "Km": "数学",
  "Z": 6,
  "L": 6
}
```

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `School` | string | `""` | 学校/任务名称 |
| `Nj` | string | `""` | 年级 |
| `ClassId` | string | `""` | 班级 |
| `Km` | string | `""` | 课程名称 |
| `Z` | int | `6` | 按钮网格行数 |
| `L` | int | `6` | 按钮网格列数 |

---

## 5. API 接口一览

| 方法 | 路径 | 用途 | 认证方式 |
|------|------|------|----------|
| GET | `/api/status` | 获取所有设备状态列表 | 无需认证 |
| GET | `/api/machines/{uuid}/tasks` | 获取设备任务列表 | 无需认证 |
| POST | `/api/register` | 客户端注册 | 密码 |
| POST | `/api/sync_data` | 客户端同步打卡数据 | 密码+RSA签名 |
| POST | `/api/load_data` | 客户端加载打卡数据 | 密码+challenge签名 |
| POST | `/api/get_config` | 客户端获取任务配置 | 密码+challenge签名 |
| POST | `/api/update_config` | 客户端更新任务配置 | 密码+RSA签名 |
| POST | `/api/update_machine_config` | Web面板修改设备配置 | 密码 |
| POST | `/api/clear_attendance` | Web面板清除打卡数据 | 密码 |
| POST | `/api/delete_machine` | Web面板删除设备 | 密码 |
| POST | `/api/web_punch` | Web面板远程打卡 | 密码 |
| POST | `/api/web_cancel_punch` | Web面板取消打卡 | 密码 |

---

## 6. 客户端 API 详解

### 6.1 获取所有设备状态

```
GET /api/status
```

**用途**：获取所有已注册设备的基本信息和在线状态。

**请求参数**：无

**响应示例**：

```json
[
  {
    "uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
    "name": "三年（1）班",
    "online": true,
    "last_seen": "2025-01-17T08:35:00.0000000+08:00",
    "task_count": 2,
    "tasks": ["default", "数学打卡"]
  }
]
```

| 响应字段 | 类型 | 说明 |
|----------|------|------|
| `uuid` | string | 设备唯一标识符 |
| `name` | string | 设备名称 |
| `online` | bool | 是否在线（5分钟内有过活动视为在线） |
| `last_seen` | string | 最后在线时间（ISO 8601） |
| `task_count` | int | 该设备拥有的任务数量 |
| `tasks` | string[] | 任务 ID 列表 |

**调用示例**：

```bash
curl http://localhost:5250/api/status
```

---

### 6.2 获取设备任务列表

```
GET /api/machines/{uuid}/tasks
```

**用途**：获取指定设备的所有打卡任务概览。

**路径参数**：

| 参数 | 类型 | 说明 |
|------|------|------|
| `uuid` | string | 设备 UUID |

**响应示例**：

```json
{
  "machine_uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "machine_name": "三年（1）班",
  "task_name": "XX实验中学",
  "tasks": [
    {
      "task_id": "default",
      "last_updated": "2025-01-17T08:35:00.0000000+08:00",
      "record_count": 15
    }
  ]
}
```

| 响应字段 | 类型 | 说明 |
|----------|------|------|
| `machine_uuid` | string | 设备 UUID |
| `machine_name` | string | 设备名称 |
| `task_name` | string | 任务显示名称（来自配置中的 School 字段） |
| `tasks[].task_id` | string | 任务 ID |
| `tasks[].last_updated` | string | 最后更新时间 |
| `tasks[].record_count` | int | 该任务的打卡记录版本数（可追溯） |

**错误响应**：

```json
{ "error": "设备不存在" }
```
> HTTP Status: 404

**调用示例**：

```bash
curl http://localhost:5250/api/machines/a1b2c3d4-e5f6-7890-abcd-ef1234567890/tasks
```

---

### 6.3 客户端注册

```
POST /api/register
```

**用途**：客户端首次启动或重装后向服务器注册，上传 RSA 公钥和设备信息，获取 UUID。若公钥已存在则返回已有 UUID（支持重装后恢复）。

**请求体**（JSON）：

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `password` | string | 是 | 服务器共享密钥 |
| `public_key` | string | 是 | 客户端 RSA 公钥（PEM 格式） |
| `name` | string | 是 | 设备名称 |
| `task_id` | string | 否 | 任务 ID，默认 `"default"` |

**请求示例**：

```json
{
  "password": "admin123",
  "public_key": "-----BEGIN RSA PUBLIC KEY-----\nMIIBCgKCAQEA...\n-----END RSA PUBLIC KEY-----",
  "name": "三年（1）班",
  "task_id": "default"
}
```

**成功响应（新注册）**：

```json
{
  "uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "existing": false
}
```
> HTTP Status: 200

**成功响应（公钥已存在，重装恢复）**：

```json
{
  "uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "existing": true
}
```
> HTTP Status: 200

| 响应字段 | 类型 | 说明 |
|----------|------|------|
| `uuid` | string | 分配或找回的设备 UUID |
| `existing` | bool | `true` 表示设备已注册过，`false` 表示新注册 |

**错误响应**：

```json
{ "error": "invalid password" }
```
> HTTP Status: 403

**调用示例**：

```bash
curl -X POST http://localhost:5250/api/register \
  -H "Content-Type: application/json" \
  -d '{"password":"admin123","public_key":"-----BEGIN RSA PUBLIC KEY-----\n...\n-----END RSA PUBLIC KEY-----","name":"三年（1）班","task_id":"default"}'
```

---

### 6.4 同步打卡数据

```
POST /api/sync_data
```

**用途**：客户端将本地的打卡数据上传到服务器。需要 RSA 签名验证，确保数据来源可信。

**请求体**（JSON）：

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `password` | string | 是 | 服务器共享密钥 |
| `uuid` | string | 是 | 设备 UUID |
| `task_id` | string | 否 | 任务 ID，默认 `"default"` |
| `data` | string | 是 | 打卡数据 JSON 字符串（`Dictionary<string, StudentAttendance>` 序列化结果） |
| `signature` | string | 是 | 对 `data` 字段内容的 RSA SHA256 签名（Base64），使用客户端私钥签名 |

**请求示例**：

```json
{
  "password": "admin123",
  "uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "task_id": "default",
  "data": "{\"张三\":{\"Name\":\"张三\",\"Count\":1,\"FirstTime\":\"2025-01-17 08:35:00\",\"History\":[\"2025-01-17 08:35:00\"]}}",
  "signature": "q2vlGx...（Base64 编码的签名）"
}
```

**成功响应**：

```json
{ "status": "ok" }
```
> HTTP Status: 200

**错误响应**：

| 错误内容 | HTTP Status | 触发条件 |
|----------|-------------|----------|
| `{"error":"invalid password"}` | 403 | 密码错误 |
| `{"error":"unknown machine"}` | 403 | UUID 对应的设备不存在 |
| `{"error":"invalid signature"}` | 403 | RSA 签名验证失败 |

**工作原理**：

1. 客户端对 `data` 字段内容（原始 JSON 字符串）使用自己的 **RSA 私钥** 计算 SHA256 签名
2. 服务器根据 UUID 查到之前注册时上传的 **RSA 公钥**
3. 服务器用该公钥验证签名，确保数据未被篡改且来自该设备
4. 验证通过后，数据以 **追加方式** 写入数据库（不覆盖历史记录，支持版本回溯）

**调用示例**：

```bash
# 假设已用 openssl 生成了签名
SIG=$(echo -n '{"张三":{"Name":"张三","Count":1,"FirstTime":"2025-01-17 08:35:00","History":["2025-01-17 08:35:00"]}}' | openssl dgst -sha256 -sign private.pem | base64 -w0)

curl -X POST http://localhost:5250/api/sync_data \
  -H "Content-Type: application/json" \
  -d "{\"password\":\"admin123\",\"uuid\":\"a1b2c3d4-...\",\"task_id\":\"default\",\"data\":\"$(echo -n '{"张三":...}' | sed 's/"/\\"/g')\",\"signature\":\"$SIG\"}"
```

---

### 6.5 加载打卡数据

```
POST /api/load_data
```

**用途**：客户端从服务器拉取最新的打卡数据（用于多端同步或恢复数据）。

**请求体**（JSON）：

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `password` | string | 是 | 服务器共享密钥 |
| `uuid` | string | 是 | 设备 UUID |
| `task_id` | string | 否 | 任务 ID，默认 `"default"` |
| `challenge` | string | 是 | 服务端生成的随机挑战字符串，客户端需对其签名 |
| `signature` | string | 是 | 对 `challenge` 字段的 RSA SHA256 签名（Base64） |

**工作流程**：

1. 客户端先请求获得一个 `challenge`（任意随机字符串，客户端自己生成）
2. 客户端用私钥对该 challenge 签名
3. 服务器验证签名通过后返回最新打卡数据

**请求示例**：

```json
{
  "password": "admin123",
  "uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "task_id": "default",
  "challenge": "random_challenge_string_12345",
  "signature": "Base64签名..."
}
```

**成功响应**：

```json
{
  "data": {
    "张三": {
      "Name": "张三",
      "Count": 3,
      "FirstTime": "2025-01-15 08:30:00",
      "History": ["2025-01-15 08:30:00", "2025-01-16 08:25:00", "2025-01-17 08:35:00"]
    }
  }
}
```

| 响应字段 | 类型 | 说明 |
|----------|------|------|
| `data` | object | 学生打卡数据字典，key 为学生姓名，value 为 `StudentAttendance` 对象 |

**错误响应**：

| 错误内容 | HTTP Status | 触发条件 |
|----------|-------------|----------|
| `{"error":"invalid password"}` | 403 | 密码错误 |
| `{"error":"unknown machine"}` | 403 | UUID 不存在 |
| `{"error":"invalid signature"}` | 403 | challenge 签名验证失败 |

**调用示例**：

```bash
CHALLENGE="random_challenge_$(date +%s)"
SIG=$(echo -n "$CHALLENGE" | openssl dgst -sha256 -sign private.pem | base64 -w0)

curl -X POST http://localhost:5250/api/load_data \
  -H "Content-Type: application/json" \
  -d "{\"password\":\"admin123\",\"uuid\":\"a1b2c3d4-...\",\"task_id\":\"default\",\"challenge\":\"$CHALLENGE\",\"signature\":\"$SIG\"}"
```

---

### 6.6 获取任务配置

```
POST /api/get_config
```

**用途**：客户端从服务器获取存储在服务器端的任务配置（学校名称、课程、网格布局等）。

**请求体**（JSON）：

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `password` | string | 是 | 服务器共享密钥 |
| `uuid` | string | 是 | 设备 UUID |
| `challenge` | string | 是 | 随机挑战字符串 |
| `signature` | string | 是 | 对 challenge 的 RSA SHA256 签名（Base64） |

**请求示例**：

```json
{
  "password": "admin123",
  "uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "challenge": "random_challenge_string_12345",
  "signature": "Base64签名..."
}
```

**成功响应**：

```json
{
  "config": {
    "School": "XX实验中学",
    "Nj": "三年级",
    "ClassId": "1班",
    "Km": "数学",
    "Z": 6,
    "L": 6
  }
}
```

| 响应字段 | 类型 | 说明 |
|----------|------|------|
| `config` | ClientConfig | 客户端任务配置对象 |

**错误响应**：同 `load_data` 接口的错误格式。

---

### 6.7 更新任务配置

```
POST /api/update_config
```

**用途**：客户端将本地的任务配置上传到服务器保存，需要 RSA 签名验证。

**请求体**（JSON）：

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `password` | string | 是 | 服务器共享密钥 |
| `uuid` | string | 是 | 设备 UUID |
| `config` | ClientConfig | 是 | 任务配置对象 |
| `signature` | string | 是 | 对 `config` 字段内容（原始 JSON）的 RSA SHA256 签名（Base64） |

**请求示例**：

```json
{
  "password": "admin123",
  "uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "config": {
    "School": "XX实验中学",
    "Nj": "三年级",
    "ClassId": "1班",
    "Km": "数学",
    "Z": 6,
    "L": 6
  },
  "signature": "Base64签名..."
}
```

**成功响应**：

```json
{ "status": "ok" }
```

**签名规则**：对 `config` 对象的 **原始 JSON 文本**（即请求中的 `"config"` 值）进行签名，而非对整个请求体签名。

---

## 7. Web 面板 API 详解

以下接口供 **Web 管理面板** 调用，认证方式为服务器密码验证（无需 RSA 签名）。

---

### 7.1 修改设备配置

```
POST /api/update_machine_config
```

**用途**：Web 管理面板中直接修改指定设备的任务配置。

**请求体**（JSON）：

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `password` | string | 是 | 服务器共享密钥 |
| `machine_uuid` | string | 是 | 设备 UUID |
| `config` | object | 是 | 新的配置 JSON 对象（完整替换） |

**请求示例**：

```json
{
  "password": "admin123",
  "machine_uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "config": { "School": "XX实验小学", "Nj": "四年级", "ClassId": "2班", "Km": "数学", "Z": 5, "L": 5 }
}
```

**成功响应**：

```json
{ "status": "ok" }
```

**错误响应**：

```json
{ "error": "invalid password" }
```
> HTTP Status: 403

---

### 7.2 清除打卡数据

```
POST /api/clear_attendance
```

**用途**：清除指定设备指定任务的打卡数据（重置为 `{}`）。

**请求体**（JSON）：

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `password` | string | 是 | 服务器共享密钥 |
| `machine_uuid` | string | 是 | 设备 UUID |
| `task_id` | string | 否 | 任务 ID，不传则清除该设备所有任务数据 |

**请求示例**：

```json
{
  "password": "admin123",
  "machine_uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "task_id": "default"
}
```

**成功响应**：

```json
{ "status": "ok" }
```

**工作原理**：

1. 删除指定设备和任务的所有打卡记录
2. 插入一条新的空记录 `Data = "{}"`，确保查询时不会报空
3. 这种方式保留了历史记录的可追溯性的同时实现了 "软清除"

---

### 7.3 删除设备

```
POST /api/delete_machine
```

**用途**：从服务器删除指定设备及其所有关联的打卡记录。

**请求体**（JSON）：

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `password` | string | 是 | 服务器共享密钥 |
| `machine_uuid` | string | 是 | 设备 UUID |

**请求示例**：

```json
{
  "password": "admin123",
  "machine_uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
}
```

**成功响应**：

```json
{ "status": "ok" }
```

**错误响应**：

```json
{ "error": "设备不存在" }
```
> HTTP Status: 404

**注意事项**：

- 此操作 **不可逆**，删除后设备的所有打卡记录均被移除
- 删除后该设备如需重新接入，需重新注册并获取新的 UUID

---

### 7.4 Web 面板远程打卡

```
POST /api/web_punch
```

**用途**：管理员在 Web 管理面板中手动为学生打卡。

**请求体**（JSON）：

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `password` | string | 是 | 服务器共享密钥 |
| `machine_uuid` | string | 是 | 设备 UUID |
| `task_id` | string | 否 | 任务 ID，默认 `"default"` |
| `student_name` | string | 是 | 学生姓名 |

**请求示例**：

```json
{
  "password": "admin123",
  "machine_uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "task_id": "default",
  "student_name": "张三"
}
```

**成功响应**：

```json
{ "status": "ok" }
```

**错误响应**：

```json
{ "error": "该学生已经打卡" }
```
> HTTP Status: 400
> 表示该学生已经打过了，不能重复打卡

**工作原理**：

1. 从数据库加载该任务最新版本的打卡数据
2. 检查该学生是否已打过卡（`FirstTime != null`），已打过则返回错误
3. 若未打过，则设置 `FirstTime`、增加 `Count`、追加 `History` 记录
4. 将修改后的数据作为新版本追加写入数据库

---

### 7.5 Web 面板取消打卡

```
POST /api/web_cancel_punch
```

**用途**：管理员在 Web 管理面板中撤销学生的打卡记录（撤回最近一次打卡）。

**请求体**（JSON）：

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `password` | string | 是 | 服务器共享密钥 |
| `machine_uuid` | string | 是 | 设备 UUID |
| `task_id` | string | 否 | 任务 ID，默认 `"default"` |
| `student_name` | string | 是 | 学生姓名 |

**请求示例**：

```json
{
  "password": "admin123",
  "machine_uuid": "a1b2c3d4-e5f6-7890-abcd-ef1234567890",
  "task_id": "default",
  "student_name": "张三"
}
```

**成功响应**：

```json
{ "status": "ok" }
```

**错误响应**：

| 错误内容 | HTTP Status | 触发条件 |
|----------|-------------|----------|
| `{"error":"该任务无打卡数据"}` | 404 | 该设备/任务没有任何打卡记录 |
| `{"error":"学生不存在"}` | 404 | 该学生不在打卡名单中 |
| `{"error":"该学生未打卡"}` | 400 | 学生 `FirstTime` 为 null，不存在可取消的打卡 |

**工作原理**：

1. 加载该任务最新版本的打卡数据
2. 找到目标学生
3. 移除 `History` 列表中的最后一条记录
4. 如果移除的正是 `FirstTime`，则将 `FirstTime` 更新为剩余历史记录的第一条（若无则置 null）
5. `Count` 减 1（最低为 0）

---

## 8. 认证页面接口

### 8.1 登录页面

```
GET /login
```

**用途**：显示 Web 管理面板的登录页面。

**响应**：`text/html` - 登录表单页面（从 `wwwroot/login.html` 模板渲染）

**调用方式**：浏览器直接访问 `http://localhost:5250/login`


### 8.2 提交登录

```
POST /login
```

**用途**：提交用户名和密码进行登录认证。

**请求格式**：HTML Form 表单（`application/x-www-form-urlencoded`）

**表单字段**：

| 字段 | 类型 | 说明 |
|------|------|------|
| `username` | string | 管理员用户名 |
| `password` | string | 管理员密码 |

**成功响应**：

- HTTP Status: 302 重定向到 `/`（首页）
- 设置 `sw_session` Cookie（HttpOnly，7天有效）

**失败响应**：

- 返回登录页面，显示错误信息"用户名或密码错误"

**调用示例**：

```bash
curl -X POST http://localhost:5250/login \
  -d "username=admin&password=admin"
```


### 8.3 退出登录

```
GET /logout
```

**用途**：清除 Session Cookie，退出登录。

**响应**：302 重定向到 `/login`

**调用方式**：浏览器访问 `http://localhost:5250/logout`

---

## 9. Web 管理面板页面

以下页面在浏览器中展示，**需要先登录**才能访问。

### 9.1 首页 - 设备总览

```
GET /
```

**用途**：显示所有已注册设备的卡片列表，包含在线/离线状态、任务数、最后在线时间。支持点击进入设备详情、删除设备操作。

**页面内容**：

- 顶部统计卡片：设备总数、在线设备数、离线设备数
- 设备列表表格：设备名称（文件夹图标）、在线状态徽章、任务数、最后在线时间、操作按钮（查看/删除）

### 9.2 设备详情页

```
GET /machine/{uuid}
```

**路径参数**：

| 参数 | 类型 | 说明 |
|------|------|------|
| `uuid` | string | 设备 UUID |

**用途**：显示指定设备的所有打卡任务（卡片式布局），每个任务显示总人数和已打卡人数。

**页面内容**：

- 面包屑导航
- 设备名称、在线状态、编辑配置按钮、删除设备按钮
- 任务卡片网格：任务图标、任务名称、总人数/已打卡统计、最后更新时间

### 9.3 任务详情页

```
GET /machine/{uuid}/task/{taskId}
```

**路径参数**：

| 参数 | 类型 | 说明 |
|------|------|------|
| `uuid` | string | 设备 UUID |
| `taskId` | string | 任务 ID（URL 编码） |

**用途**：显示打卡排名和打卡状态网格，支持 Web 远程打卡/取消打卡操作。

**页面内容**：

- 面包屑导航
- 统计卡片：总人数、已打卡、未打卡
- 打卡排名表：排名、姓名、打卡时间
- 学生打卡状态网格：点击可进行打卡或取消打卡操作
- 清除数据按钮

---

## 10. 调用示例

### 10.1 完整的客户端工作流

```
┌─────────┐                              ┌─────────┐
│ Client  │                              │ Server  │
└────┬────┘                              └────┬────┘
     │                                        │
     │  1. POST /api/register                  │
     │     (password + public_key + name)      │
     │ ──────────────────────────────────────> │
     │                                        │
     │     返回 { uuid, existing }             │
     │ <────────────────────────────────────── │
     │                                        │
     │  2. POST /api/get_config                │
     │     (password + uuid + challenge + sig) │
     │ ──────────────────────────────────────> │
     │                                        │
     │     返回 { config }                      │
     │ <────────────────────────────────────── │
     │                                        │
     │  3. POST /api/sync_data                 │
     │     (password + uuid + data + sig)      │
     │ ──────────────────────────────────────> │
     │                                        │
     │     返回 { status: "ok" }               │
     │ <────────────────────────────────────── │
     │                                        │
     │  4. POST /api/load_data                 │
     │     (password + uuid + challenge + sig) │
     │ ──────────────────────────────────────> │
     │                                        │
     │     返回 { data: {...} }                │
     │ <────────────────────────────────────── │
     │                                        │
```

### 10.2 使用 PowerShell 调用

```powershell
$ServerUrl = "http://localhost:5250"
$Password  = "admin123"

# 1. 查看所有设备
Invoke-RestMethod -Uri "$ServerUrl/api/status" | ConvertTo-Json

# 2. 注册设备（假设已有公钥）
$Body = @{
    password   = $Password
    public_key = "-----BEGIN RSA PUBLIC KEY-----..."
    name       = "测试设备"
    task_id    = "default"
} | ConvertTo-Json

$Result = Invoke-RestMethod -Uri "$ServerUrl/api/register" -Method Post -Body $Body -ContentType "application/json"
$Uuid = $Result.uuid
Write-Host "注册成功，UUID: $Uuid"

# 3. 查看设备任务
Invoke-RestMethod -Uri "$ServerUrl/api/machines/$Uuid/tasks" | ConvertTo-Json
```

### 10.3 使用 Python 调用

```python
import requests
import json

BASE_URL = "http://localhost:5250"
PASSWORD = "admin123"

# 1. 获取设备列表
resp = requests.get(f"{BASE_URL}/api/status")
print(json.dumps(resp.json(), indent=2, ensure_ascii=False))

# 2. 注册设备
resp = requests.post(f"{BASE_URL}/api/register", json={
    "password": PASSWORD,
    "public_key": "-----BEGIN RSA PUBLIC KEY-----...",
    "name": "测试设备"
})
result = resp.json()
uuid = result["uuid"]
print(f"UUID: {uuid}, 是否已有: {result['existing']}")

# 3. Web 面板远程打卡
resp = requests.post(f"{BASE_URL}/api/web_punch", json={
    "password": PASSWORD,
    "machine_uuid": uuid,
    "task_id": "default",
    "student_name": "张三"
})
print(resp.json())
```

---

> **文档版本**：对应 CheckIn.Net V2.6 分支
> **最后更新**：2026-07-11
