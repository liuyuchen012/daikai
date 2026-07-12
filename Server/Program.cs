using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using CheckIn.Server.Data;
using CheckIn.Shared.Models;
using Microsoft.EntityFrameworkCore;
using System.Web;

// =============================================================================
// AgoraIn 集控平台 - 服务器端入口
// 功能：设备注册与管理、打卡数据同步、Web 管理面板（含多用户登录认证和权限管理）
//       移动端 API（JWT 令牌认证、二维码签到、管理员仪表盘）
// =============================================================================

// ---- SHA256 哈希辅助方法 ----
string Sha256(string input)
{
    var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(input));
    return Convert.ToHexString(bytes).ToLower();
}

// ---- 加载服务器配置文件 config.json ----
var configPath = Path.Combine(AppContext.BaseDirectory, "config.json");
if (!File.Exists(configPath))
    configPath = Path.Combine(Directory.GetCurrentDirectory(), "config.json");

var configJson = File.Exists(configPath) ? JsonDocument.Parse(File.ReadAllText(configPath)).RootElement : default;
var cfgPort = configJson.TryGetProperty("Port", out var p) ? p.GetInt32() : 5250;
var serverName = configJson.TryGetProperty("ServerName", out var sn) ? sn.GetString() ?? "AgoraIn 集控平台" : "AgoraIn 集控平台";
var serverPassword = configJson.TryGetProperty("ServerPassword", out var sp) ? sp.GetString() ?? "admin123" : "admin123";

var builder = WebApplication.CreateBuilder(args);
builder.WebHost.UseUrls($"http://0.0.0.0:{cfgPort}");

var connStr = builder.Configuration.GetConnectionString("Default") ?? "Data Source=checkin.db";
builder.Services.AddDbContext<AppDbContext>(o => o.UseSqlite(connStr));
builder.Services.AddCors(o => o.AddDefaultPolicy(p => p.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader()));

var app = builder.Build();
app.UseCors();

using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();

    // ---- 从 config.json 种子用户数据（仅在 Users 表为空时） ----
    if (!db.Users.Any() && configJson.TryGetProperty("Users", out var usersArr))
    {
        foreach (var userEl in usersArr.EnumerateArray())
        {
            var username = userEl.TryGetProperty("Username", out var un) ? un.GetString() ?? "" : "";
            var passwordHash = userEl.TryGetProperty("PasswordHash", out var ph) ? ph.GetString() ?? "" : "";
            var role = userEl.TryGetProperty("Role", out var r) ? r.GetString() ?? "viewer" : "viewer";
            var displayName = userEl.TryGetProperty("DisplayName", out var dn) ? dn.GetString() ?? "" : "";
            var isActive = userEl.TryGetProperty("IsActive", out var ia) && ia.GetBoolean();

            if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(passwordHash))
            {
                db.Users.Add(new UserEntity
                {
                    Username = username,
                    PasswordHash = passwordHash,
                    Role = role,
                    DisplayName = displayName,
                    IsActive = isActive,
                    CreatedAt = DateTime.Now.ToString("O")
                });
            }
        }
        db.SaveChanges();
    }
}

// ---- 会话认证（Session Auth） ----
// 使用 HMAC-SHA256 签名的 Cookie 实现无状态会话管理
var sessionSecret = Guid.NewGuid().ToString("N");
const string CookieName = "sw_session";

/// <summary>
/// 创建带签名的会话令牌，格式：username:role:timestamp:signature
/// </summary>
string MakeToken(string username, string role = "viewer")
{
    var ts = DateTimeOffset.UtcNow.ToUnixTimeSeconds().ToString();
    var payload = $"{username}:{role}:{ts}";
    using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(sessionSecret));
    var sig = Convert.ToHexString(hmac.ComputeHash(Encoding.UTF8.GetBytes(payload))).ToLower();
    return $"{payload}:{sig}";
}

/// <summary>
/// 验证会话令牌的签名是否有效（防篡改）
/// </summary>
bool ValidateToken(string token)
{
    var parts = token.Split(':');
    if (parts.Length != 4) return false;
    var payload = $"{parts[0]}:{parts[1]}:{parts[2]}";
    using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(sessionSecret));
    var expected = Convert.ToHexString(hmac.ComputeHash(Encoding.UTF8.GetBytes(payload))).ToLower();
    return expected == parts[3];
}

/// <summary>
/// 检查当前请求是否已认证（通过 Cookie 中的会话令牌）
/// </summary>
bool IsAuthenticated(HttpContext ctx)
{
    if (!ctx.Request.Cookies.TryGetValue(CookieName, out var token)) return false;
    return ValidateToken(token);
}

/// <summary>
/// 从会话令牌中提取用户名
/// </summary>
string? GetUsername(HttpContext ctx)
{
    if (!ctx.Request.Cookies.TryGetValue(CookieName, out var token)) return null;
    if (!ValidateToken(token)) return null;
    var parts = token.Split(':');
    return parts.Length >= 2 ? parts[0] : null;
}

/// <summary>
/// 从会话令牌中提取用户角色
/// </summary>
string? GetUserRole(HttpContext ctx)
{
    if (!ctx.Request.Cookies.TryGetValue(CookieName, out var token)) return null;
    if (!ValidateToken(token)) return null;
    var parts = token.Split(':');
    return parts.Length >= 3 ? parts[1] : null;
}

/// <summary>
/// 检查用户是否拥有指定角色之一
/// </summary>
bool HasRole(HttpContext ctx, string[] roles)
{
    var role = GetUserRole(ctx);
    return role != null && roles.Contains(role);
}

/// <summary>
/// 检查用户是否为管理员
/// </summary>
bool IsAdmin(HttpContext ctx) => HasRole(ctx, new[] { "admin" });

/// <summary>
/// 检查用户是否为管理员或操作员
/// </summary>
bool IsAdminOrOperator(HttpContext ctx) => HasRole(ctx, new[] { "admin", "operator" });

// ---- Bearer Token 认证（用于移动端 API） ----
// 复用现有的 HMAC Token 机制，但通过 Authorization: Bearer <token> 头传递

/// <summary>
/// 从请求的 Authorization 头中提取并验证 Bearer Token
/// </summary>
(string? Username, string? Role, string? Error) ParseBearerToken(HttpContext ctx)
{
    var authHeader = ctx.Request.Headers.Authorization.FirstOrDefault();
    if (string.IsNullOrEmpty(authHeader) || !authHeader.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
        return (null, null, "缺少认证令牌");

    var token = authHeader["Bearer ".Length..].Trim();
    if (!ValidateToken(token))
        return (null, null, "令牌无效或已过期");

    var parts = token.Split(':');
    if (parts.Length < 3)
        return (null, null, "令牌格式错误");

    // 检查过期（7天有效期）
    if (long.TryParse(parts[2], out var ts))
    {
        var tokenTime = DateTimeOffset.FromUnixTimeSeconds(ts);
        if ((DateTimeOffset.UtcNow - tokenTime).TotalDays > 7)
            return (null, null, "令牌已过期");
    }

    return (parts[0], parts[1], null);
}

/// <summary>
/// 通过 Bearer Token 获取用户名
/// </summary>
string? GetBearerUsername(HttpContext ctx)
{
    var (username, _, error) = ParseBearerToken(ctx);
    return error == null ? username : null;
}

/// <summary>
/// 通过 Bearer Token 获取用户角色
/// </summary>
string? GetBearerUserRole(HttpContext ctx)
{
    var (_, role, error) = ParseBearerToken(ctx);
    return error == null ? role : null;
}

/// <summary>
/// 通过 Bearer Token 检查是否为管理员
/// </summary>
bool IsBearerAdmin(HttpContext ctx)
{
    var role = GetBearerUserRole(ctx);
    return role == "admin";
}

/// <summary>
/// 生成移动端登录令牌（与 Web Cookie 令牌格式相同，但作为 JSON 返回）
/// </summary>
string MakeLoginToken(string username, string role)
{
    return MakeToken(username, role);
}

// ---- 加载 HTML 模板文件（template.html 和 login.html） ----
var wwwroot = builder.Environment.WebRootPath ?? "wwwroot";
var templatePath = Path.Combine(wwwroot, "template.html");
var loginTemplatePath = Path.Combine(wwwroot, "login.html");
var templateContent = File.Exists(templatePath) ? File.ReadAllText(templatePath) : "<html><body>{CONTENT}</body></html>";
var loginTemplateContent = File.Exists(loginTemplatePath) ? File.ReadAllText(loginTemplatePath) : "<html><body>Login</body></html>";

/// <summary>
/// 将页面内容嵌入模板，替换标题和导航高亮
/// </summary>
string RenderPage(string content, string activeNav = "home", HttpContext? ctx = null)
{
    var isAdminUser = ctx != null && IsAdmin(ctx);
    var username = ctx != null ? GetUsername(ctx) ?? "" : "";

    return templateContent
        .Replace("{TITLE}", HttpUtility.HtmlEncode(serverName))
        .Replace("{NAV_HOME}", activeNav == "home" ? "active" : "")
        .Replace("{NAV_USERS}", activeNav == "users" ? "active" : "")
        .Replace("{NAV_PROFILE}", activeNav == "profile" ? "active" : "")
        .Replace("{USERS_VISIBLE}", isAdminUser ? "block" : "none")
        .Replace("{CURRENT_USER}", HttpUtility.HtmlEncode(username))
        .Replace("{CONTENT}", content);
}

/// <summary>
/// 渲染登录页面，支持显示错误消息
/// </summary>
string RenderLoginPage(string errorMsg = "")
{
    var hasErr = !string.IsNullOrEmpty(errorMsg);
    return loginTemplateContent
        .Replace("{TITLE}", HttpUtility.HtmlEncode(serverName))
        .Replace("{ERROR_DISPLAY}", hasErr ? "block" : "none")
        .Replace("{ERROR_MSG}", HttpUtility.HtmlEncode(errorMsg));
}

// ---- 认证中间件：保护所有页面路由，跳过 /api/*、/login、/logout 和 /static ----
app.Use(async (ctx, next) =>
{
    var path = ctx.Request.Path.Value ?? "/";
    // Allow: login page, logout, API endpoints, static files, sign-in pages
    if (path.StartsWith("/api/") || path == "/login" || path == "/logout" || path.StartsWith("/static") || path.StartsWith("/s/"))
    {
        await next();
        return;
    }
    // Check auth for all other pages
    if (!IsAuthenticated(ctx))
    {
        ctx.Response.Redirect("/login");
        return;
    }
    await next();
});

// ---- RSA 签名验证辅助方法 ----
/// <summary>
/// 使用客户端公钥验证签名（RSA PKCS1 SHA256）
/// </summary>
bool VerifySignature(string pubKeyPem, string message, string signatureB64)
{
    try
    {
        var rsa = RSA.Create();
        rsa.ImportFromPem(pubKeyPem);
        return rsa.VerifyData(Encoding.UTF8.GetBytes(message), Convert.FromBase64String(signatureB64), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
    }
    catch { return false; }
}

/// <summary>
/// 从数据库获取指定设备的 RSA 公钥
/// </summary>
string? GetPublicKey(AppDbContext db, string uuid) =>
    db.Machines.Where(m => m.Uuid == uuid).Select(m => m.PublicKey).FirstOrDefault();

/// <summary>
/// 验证请求中的密码是否与服务器密码一致
/// </summary>
static bool CheckPwd(JsonElement body, string expected)
{
    if (!body.TryGetProperty("password", out var p) || p.ValueKind != JsonValueKind.String)
        return false;
    return p.GetString() == expected;
}

// ===== API 端点 =====

/// <summary>
/// GET /api/status - 获取所有已注册设备列表（含在线状态和任务数量）
/// </summary>
app.MapGet("/api/status", async (AppDbContext db) =>
{
    var machines = await db.Machines.ToListAsync();
    var attendances = await db.AttendanceRecords.ToListAsync();
    var now = DateTime.Now;

    // 按设备分组
    var groupedMachines = machines.Select(m => {
        var tasks = attendances.Where(a => a.MachineUuid == m.Uuid).Select(a => a.TaskId).Distinct().ToList();
        var last = m.LastSeen != null ? DateTime.Parse(m.LastSeen) : (DateTime?)null;
        var online = last != null && (now - last.Value).TotalSeconds < 300;

        return new {
            uuid = m.Uuid,
            name = m.Name,
            online,
            last_seen = m.LastSeen,
            task_count = tasks.Count,
            tasks
        };
    }).ToList();

    return Results.Json(groupedMachines);
});

/// <summary>
/// GET /api/machines/{uuid}/tasks - 获取指定设备的所有任务列表
/// </summary>
app.MapGet("/api/machines/{uuid}/tasks", async (string uuid, AppDbContext db) =>
{
    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Json(new { error = "设备不存在" }, statusCode: 404);

    var tasks = await db.AttendanceRecords
        .Where(a => a.MachineUuid == uuid)
        .GroupBy(a => a.TaskId)
        .Select(g => new {
            task_id = g.Key,
            last_updated = g.Max(a => a.UpdatedAt),
            record_count = g.Count()
        })
        .ToListAsync();

    var machineConfig = JsonSerializer.Deserialize<ClientConfig>(machine.Config) ?? new ClientConfig();
    var taskName = machineConfig.School ?? machine.Name;

    return Results.Json(new {
        machine_uuid = uuid,
        machine_name = machine.Name,
        task_name = taskName,
        tasks
    });
});

/// <summary>
/// POST /api/register - 客户端注册：上传公钥和设备信息，获取 UUID
/// </summary>
app.MapPost("/api/register", async (AppDbContext db, JsonElement body) =>
{
    if (!CheckPwd(body, serverPassword))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var pubKey = body.GetProperty("public_key").GetString() ?? "";
    var name = body.GetProperty("name").GetString() ?? "";
    var taskId = body.TryGetProperty("task_id", out var tid) ? tid.GetString() ?? "default" : "default";

    var existing = await db.Machines.FirstOrDefaultAsync(m => m.PublicKey == pubKey);
    if (existing != null)
    {
        existing.LastSeen = DateTime.Now.ToString("O");
        await db.SaveChangesAsync();
        return Results.Json(new { uuid = existing.Uuid, existing = true });
    }

    var machine = new MachineEntity
    {
        Uuid = Guid.NewGuid().ToString(),
        Name = name,
        PublicKey = pubKey,
        LastSeen = DateTime.Now.ToString("O")
    };
    db.Machines.Add(machine);
    await db.SaveChangesAsync();
    return Results.Json(new { uuid = machine.Uuid, existing = false });
});

/// <summary>
/// POST /api/sync_data - 客户端同步打卡数据到服务器（需 RSA 签名验证）
/// </summary>
app.MapPost("/api/sync_data", async (AppDbContext db, JsonElement body) =>
{
    if (!CheckPwd(body, serverPassword))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("uuid").GetString() ?? "";
    var taskId = body.TryGetProperty("task_id", out var tid) ? tid.GetString() ?? "default" : "default";
    var signature = body.GetProperty("signature").GetString() ?? "";
    var data = body.GetProperty("data").GetString() ?? "";
    var pubKey = GetPublicKey(db, uuid);
    if (pubKey == null) return Results.Json(new { error = "unknown machine" }, statusCode: 403);
    if (!VerifySignature(pubKey, data, signature))
        return Results.Json(new { error = "invalid signature" }, statusCode: 403);

    db.AttendanceRecords.Add(new AttendanceEntity {
        MachineUuid = uuid,
        TaskId = taskId,
        Data = data,
        UpdatedAt = DateTime.Now.ToString("O")
    });
    var m = await db.Machines.FindAsync(uuid);
    if (m != null) m.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

/// <summary>
/// POST /api/load_data - 客户端从服务器加载最新打卡数据（需 challenge-签名验证）
/// </summary>
app.MapPost("/api/load_data", async (AppDbContext db, JsonElement body) =>
{
    if (!CheckPwd(body, serverPassword))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("uuid").GetString() ?? "";
    var taskId = body.TryGetProperty("task_id", out var tid) ? tid.GetString() ?? "default" : "default";
    var signature = body.GetProperty("signature").GetString() ?? "";
    var challenge = body.GetProperty("challenge").GetString() ?? "";
    var pubKey = GetPublicKey(db, uuid);
    if (pubKey == null) return Results.Json(new { error = "unknown machine" }, statusCode: 403);
    if (!VerifySignature(pubKey, challenge, signature))
        return Results.Json(new { error = "invalid signature" }, statusCode: 403);

    var latest = await db.AttendanceRecords
        .Where(a => a.MachineUuid == uuid && a.TaskId == taskId)
        .OrderByDescending(a => a.UpdatedAt)
        .FirstOrDefaultAsync();

    var m = await db.Machines.FindAsync(uuid);
    if (m != null) m.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();

    var data = latest != null ? JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latest.Data) : new();
    return Results.Json(new { data });
});

/// <summary>
/// POST /api/get_config - 客户端获取存储在服务器上的任务配置
/// </summary>
app.MapPost("/api/get_config", async (AppDbContext db, JsonElement body) =>
{
    if (!CheckPwd(body, serverPassword))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("uuid").GetString() ?? "";
    var signature = body.GetProperty("signature").GetString() ?? "";
    var challenge = body.GetProperty("challenge").GetString() ?? "";
    var pubKey = GetPublicKey(db, uuid);
    if (pubKey == null) return Results.Json(new { error = "unknown machine" }, statusCode: 403);
    if (!VerifySignature(pubKey, challenge, signature))
        return Results.Json(new { error = "invalid signature" }, statusCode: 403);

    var machine = await db.Machines.FindAsync(uuid);
    if (machine != null) machine.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();

    var config = machine != null ? JsonSerializer.Deserialize<ClientConfig>(machine.Config) ?? new ClientConfig() : new ClientConfig();
    return Results.Json(new { config });
});

/// <summary>
/// POST /api/update_config - 客户端更新存储在服务器上的任务配置
/// </summary>
app.MapPost("/api/update_config", async (AppDbContext db, JsonElement body) =>
{
    if (!CheckPwd(body, serverPassword))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("uuid").GetString() ?? "";
    var signature = body.GetProperty("signature").GetString() ?? "";
    var configStr = body.GetProperty("config").GetRawText();
    var pubKey = GetPublicKey(db, uuid);
    if (pubKey == null) return Results.Json(new { error = "unknown machine" }, statusCode: 403);
    if (!VerifySignature(pubKey, configStr, signature))
        return Results.Json(new { error = "invalid signature" }, statusCode: 403);

    var machine = await db.Machines.FindAsync(uuid);
    if (machine != null) { machine.Config = configStr; machine.LastSeen = DateTime.Now.ToString("O"); await db.SaveChangesAsync(); }
    return Results.Json(new { status = "ok" });
});

/// <summary>
/// POST /api/update_machine_config - Web 面板修改设备配置（需服务器密码）
/// </summary>
app.MapPost("/api/update_machine_config", async (AppDbContext db, JsonElement body) =>
{
    var pwd = body.GetProperty("password").GetString() ?? "";
    if (pwd != serverPassword) return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("machine_uuid").GetString() ?? "";
    var machine = await db.Machines.FindAsync(uuid);
    if (machine != null) { machine.Config = body.GetProperty("config").GetRawText(); await db.SaveChangesAsync(); }
    return Results.Json(new { status = "ok" });
});

/// <summary>
/// POST /api/clear_attendance - Web 面板清除指定设备/任务的打卡数据
/// </summary>
app.MapPost("/api/clear_attendance", async (AppDbContext db, JsonElement body) =>
{
    var pwd = body.GetProperty("password").GetString() ?? "";
    if (pwd != serverPassword) return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("machine_uuid").GetString() ?? "";
    var taskId = body.TryGetProperty("task_id", out var tid) ? tid.GetString() ?? null : null;

    var query = db.AttendanceRecords.Where(a => a.MachineUuid == uuid);
    if (taskId != null) query = query.Where(a => a.TaskId == taskId);

    db.AttendanceRecords.RemoveRange(query);
    db.AttendanceRecords.Add(new AttendanceEntity {
        MachineUuid = uuid,
        TaskId = taskId ?? "default",
        Data = "{}",
        UpdatedAt = DateTime.Now.ToString("O")
    });
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

/// <summary>
/// POST /api/delete_machine - Web 面板删除设备及其所有打卡数据
/// </summary>
app.MapPost("/api/delete_machine", async (AppDbContext db, JsonElement body) =>
{
    var pwd = body.GetProperty("password").GetString() ?? "";
    if (pwd != serverPassword) return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("machine_uuid").GetString() ?? "";
    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Json(new { error = "设备不存在" }, statusCode: 404);

    db.AttendanceRecords.RemoveRange(db.AttendanceRecords.Where(a => a.MachineUuid == uuid));
    db.Machines.Remove(machine);
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

/// <summary>
/// POST /api/web_punch - Web 面板远程为学生打卡
/// </summary>
app.MapPost("/api/web_punch", async (AppDbContext db, JsonElement body) =>
{
    var pwd = body.GetProperty("password").GetString() ?? "";
    if (pwd != serverPassword) return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("machine_uuid").GetString() ?? "";
    var taskId = body.TryGetProperty("task_id", out var tid) ? tid.GetString() ?? "default" : "default";
    var student = body.GetProperty("student_name").GetString() ?? "";

    var latest = await db.AttendanceRecords
        .Where(a => a.MachineUuid == uuid && a.TaskId == taskId)
        .OrderByDescending(a => a.UpdatedAt)
        .FirstOrDefaultAsync();

    var data = latest != null ? JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latest.Data) ?? new() : new();

    if (!data.ContainsKey(student)) data[student] = new StudentAttendance { Name = student };
    if (data[student].FirstTime != null) return Results.Json(new { error = "该学生已经打卡" }, statusCode: 400);

    var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    data[student].FirstTime = now; data[student].Count++; data[student].History.Add(now);

    db.AttendanceRecords.Add(new AttendanceEntity {
        MachineUuid = uuid,
        TaskId = taskId,
        Data = JsonSerializer.Serialize(data),
        UpdatedAt = DateTime.Now.ToString("O")
    });
    var m = await db.Machines.FindAsync(uuid);
    if (m != null) m.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

/// <summary>
/// POST /api/web_cancel_punch - Web 面板取消学生打卡
/// </summary>
app.MapPost("/api/web_cancel_punch", async (AppDbContext db, JsonElement body) =>
{
    var pwd = body.GetProperty("password").GetString() ?? "";
    if (pwd != serverPassword) return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("machine_uuid").GetString() ?? "";
    var taskId = body.TryGetProperty("task_id", out var tid) ? tid.GetString() ?? "default" : "default";
    var student = body.GetProperty("student_name").GetString() ?? "";

    var latest = await db.AttendanceRecords
        .Where(a => a.MachineUuid == uuid && a.TaskId == taskId)
        .OrderByDescending(a => a.UpdatedAt)
        .FirstOrDefaultAsync();

    if (latest == null) return Results.Json(new { error = "该任务无打卡数据" }, statusCode: 404);

    var data = JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latest.Data) ?? new();
    if (!data.ContainsKey(student)) return Results.Json(new { error = "学生不存在" }, statusCode: 404);
    if (data[student].FirstTime == null) return Results.Json(new { error = "该学生未打卡" }, statusCode: 400);

    var sa = data[student];
    if (sa.History.Count > 0)
    {
        var r = sa.History.Last(); sa.History.RemoveAt(sa.History.Count - 1);
        if (sa.FirstTime == r) sa.FirstTime = sa.History.Count > 0 ? sa.History.First() : null;
        sa.Count = Math.Max(0, sa.Count - 1);
    }
    else { sa.FirstTime = null; sa.Count = 0; }

    db.AttendanceRecords.Add(new AttendanceEntity {
        MachineUuid = uuid,
        TaskId = taskId,
        Data = JsonSerializer.Serialize(data),
        UpdatedAt = DateTime.Now.ToString("O")
    });
    var m = await db.Machines.FindAsync(uuid);
    if (m != null) m.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

// ===== 用户管理 API =====

/// <summary>
/// GET /api/users - 列出所有用户（需要管理员角色）
/// </summary>
app.MapGet("/api/users", async (AppDbContext db, HttpContext ctx) =>
{
    if (!IsAdmin(ctx))
        return Results.Json(new { error = "权限不足，仅管理员可访问" }, statusCode: 403);

    var users = await db.Users
        .OrderBy(u => u.Id)
        .Select(u => new
        {
            id = u.Id,
            username = u.Username,
            role = u.Role,
            display_name = u.DisplayName,
            created_at = u.CreatedAt,
            is_active = u.IsActive
        })
        .ToListAsync();

    return Results.Json(users);
});

/// <summary>
/// POST /api/users - 创建用户（需要管理员角色）
/// </summary>
app.MapPost("/api/users", async (AppDbContext db, HttpContext ctx) =>
{
    if (!IsAdmin(ctx))
        return Results.Json(new { error = "权限不足，仅管理员可创建用户" }, statusCode: 403);

    string bodyStr;
    using (var reader = new StreamReader(ctx.Request.Body, Encoding.UTF8))
    {
        bodyStr = await reader.ReadToEndAsync();
    }

    JsonElement body;
    try { body = JsonDocument.Parse(bodyStr).RootElement; }
    catch { return Results.Json(new { error = "无效的 JSON 格式" }, statusCode: 400); }

    var username = body.TryGetProperty("username", out var un) ? un.GetString()?.Trim() ?? "" : "";
    var password = body.TryGetProperty("password", out var pw) ? pw.GetString() ?? "" : "";
    var role = body.TryGetProperty("role", out var r) ? r.GetString() ?? "viewer" : "viewer";
    var displayName = body.TryGetProperty("display_name", out var dn) ? dn.GetString()?.Trim() ?? "" : "";

    if (string.IsNullOrEmpty(username))
        return Results.Json(new { error = "用户名不能为空" }, statusCode: 400);
    if (string.IsNullOrEmpty(password))
        return Results.Json(new { error = "密码不能为空" }, statusCode: 400);
    if (!new[] { "admin", "operator", "viewer" }.Contains(role))
        return Results.Json(new { error = "无效的角色，必须是 admin、operator 或 viewer" }, statusCode: 400);

    if (await db.Users.AnyAsync(u => u.Username == username))
        return Results.Json(new { error = "用户名已存在" }, statusCode: 409);

    var user = new UserEntity
    {
        Username = username,
        PasswordHash = Sha256(password),
        Role = role,
        DisplayName = string.IsNullOrEmpty(displayName) ? username : displayName,
        IsActive = true,
        CreatedAt = DateTime.Now.ToString("O")
    };
    db.Users.Add(user);
    await db.SaveChangesAsync();

    return Results.Json(new
    {
        id = user.Id,
        username = user.Username,
        role = user.Role,
        display_name = user.DisplayName,
        created_at = user.CreatedAt,
        is_active = user.IsActive
    });
});

/// <summary>
/// PUT /api/users/{id} - 更新用户信息（管理员或本人）
/// </summary>
app.MapPut("/api/users/{id}", async (int id, AppDbContext db, HttpContext ctx) =>
{
    var currentUsername = GetUsername(ctx);
    if (currentUsername == null)
        return Results.Json(new { error = "未认证" }, statusCode: 401);

    var user = await db.Users.FindAsync(id);
    if (user == null)
        return Results.Json(new { error = "用户不存在" }, statusCode: 404);

    // 只有管理员或用户本人可以更新
    if (!IsAdmin(ctx) && user.Username != currentUsername)
        return Results.Json(new { error = "权限不足，只能修改自己的信息" }, statusCode: 403);

    string bodyStr;
    using (var reader = new StreamReader(ctx.Request.Body, Encoding.UTF8))
    {
        bodyStr = await reader.ReadToEndAsync();
    }

    JsonElement body;
    try { body = JsonDocument.Parse(bodyStr).RootElement; }
    catch { return Results.Json(new { error = "无效的 JSON 格式" }, statusCode: 400); }

    // 非管理员只能修改 display_name
    if (IsAdmin(ctx))
    {
        if (body.TryGetProperty("role", out var r))
        {
            var role = r.GetString() ?? "viewer";
            if (new[] { "admin", "operator", "viewer" }.Contains(role))
                user.Role = role;
        }
        if (body.TryGetProperty("is_active", out var ia))
            user.IsActive = ia.GetBoolean();
    }

    if (body.TryGetProperty("display_name", out var dn))
    {
        var displayName = dn.GetString()?.Trim();
        if (!string.IsNullOrEmpty(displayName))
            user.DisplayName = displayName;
    }

    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

/// <summary>
/// DELETE /api/users/{id} - 删除用户（需要管理员角色，不能删除自己）
/// </summary>
app.MapDelete("/api/users/{id}", async (int id, AppDbContext db, HttpContext ctx) =>
{
    if (!IsAdmin(ctx))
        return Results.Json(new { error = "权限不足，仅管理员可删除用户" }, statusCode: 403);

    var currentUsername = GetUsername(ctx);
    var user = await db.Users.FindAsync(id);
    if (user == null)
        return Results.Json(new { error = "用户不存在" }, statusCode: 404);

    if (user.Username == currentUsername)
        return Results.Json(new { error = "不能删除自己的账户" }, statusCode: 400);

    db.Users.Remove(user);
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

/// <summary>
/// POST /api/users/change-password - 修改密码（本人或管理员）
/// </summary>
app.MapPost("/api/users/change-password", async (AppDbContext db, HttpContext ctx) =>
{
    var currentUsername = GetUsername(ctx);
    if (currentUsername == null)
        return Results.Json(new { error = "未认证" }, statusCode: 401);

    string bodyStr;
    using (var reader = new StreamReader(ctx.Request.Body, Encoding.UTF8))
    {
        bodyStr = await reader.ReadToEndAsync();
    }

    JsonElement body;
    try { body = JsonDocument.Parse(bodyStr).RootElement; }
    catch { return Results.Json(new { error = "无效的 JSON 格式" }, statusCode: 400); }

    var oldPassword = body.TryGetProperty("old_password", out var op) ? op.GetString() ?? "" : "";
    var newPassword = body.TryGetProperty("new_password", out var np) ? np.GetString() ?? "" : "";

    if (string.IsNullOrEmpty(newPassword))
        return Results.Json(new { error = "新密码不能为空" }, statusCode: 400);

    // 管理员可以通过 user_id 指定目标用户，否则修改自己
    UserEntity? targetUser;
    if (IsAdmin(ctx) && body.TryGetProperty("user_id", out var uid) && uid.ValueKind == JsonValueKind.Number)
    {
        targetUser = await db.Users.FindAsync(uid.GetInt32());
        if (targetUser == null)
            return Results.Json(new { error = "目标用户不存在" }, statusCode: 404);
    }
    else
    {
        targetUser = await db.Users.FirstOrDefaultAsync(u => u.Username == currentUsername);
        if (targetUser == null)
            return Results.Json(new { error = "用户不存在" }, statusCode: 404);
    }

    // 非管理员修改自己时必须提供旧密码
    if (!IsAdmin(ctx) || targetUser.Username == currentUsername)
    {
        if (string.IsNullOrEmpty(oldPassword))
            return Results.Json(new { error = "旧密码不能为空" }, statusCode: 400);
        if (Sha256(oldPassword) != targetUser.PasswordHash)
            return Results.Json(new { error = "旧密码错误" }, statusCode: 403);
    }

    targetUser.PasswordHash = Sha256(newPassword);
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

// ===== 签到任务 API =====

/// <summary>
/// 生成 6 位随机短链码（字母数字混合）
/// </summary>
string GenerateShortCode()
{
    const string chars = "abcdefghijkmnpqrstuvwxyz23456789"; // 去掉容易混淆的 0/O/1/l
    var bytes = RandomNumberGenerator.GetBytes(6);
    var sb = new StringBuilder(6);
    foreach (var b in bytes) sb.Append(chars[b % chars.Length]);
    return sb.ToString();
}

/// <summary>
/// POST /api/create_signin - 教师客户端创建签到任务，返回短链供学生访问
/// </summary>
app.MapPost("/api/create_signin", async (AppDbContext db, JsonElement body) =>
{
    if (!CheckPwd(body, serverPassword))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("uuid").GetString() ?? "";
    var signPassword = body.GetProperty("sign_password").GetString() ?? "";
    var classroom = body.GetProperty("classroom").GetString() ?? "";
    var subject = body.GetProperty("subject").GetString() ?? "";
    var studentList = body.TryGetProperty("students", out var sl) ? sl.GetRawText() : "[]";

    // 验证设备存在
    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Json(new { error = "unknown machine" }, statusCode: 403);

    // 生成唯一短链码（确保不重复）
    string shortCode;
    do { shortCode = GenerateShortCode(); }
    while (await db.SignInTasks.AnyAsync(s => s.ShortCode == shortCode));

    var task = new SignInTaskEntity
    {
        ShortCode = shortCode,
        MachineUuid = uuid,
        Password = signPassword,
        Classroom = classroom,
        Subject = subject,
        StudentList = studentList,
        SignInRecords = "[]",
        CreatedAt = DateTime.Now.ToString("O"),
        Status = "active"
    };
    db.SignInTasks.Add(task);
    if (machine != null) machine.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();

    // 同时在 attendance 表中创建对应任务记录（初始化为学生名单字典）
    var taskId = $"signin_{shortCode}";
    var studentNames = JsonSerializer.Deserialize<List<string>>(studentList) ?? new();
    var initialData = new Dictionary<string, StudentAttendance>();
    foreach (var name in studentNames)
        initialData[name] = new StudentAttendance { Name = name };

    db.AttendanceRecords.Add(new AttendanceEntity
    {
        MachineUuid = uuid,
        TaskId = taskId,
        Data = JsonSerializer.Serialize(initialData),
        UpdatedAt = DateTime.Now.ToString("O")
    });
    await db.SaveChangesAsync();

    return Results.Json(new
    {
        short_code = shortCode,
        task_id = taskId
    });
});

/// <summary>
/// POST /api/signin_result - 客户端拉取签到结果（含 challenge-签名验证）
/// 返回该设备下所有活跃签到任务的签到记录
/// </summary>
app.MapPost("/api/signin_result", async (AppDbContext db, JsonElement body) =>
{
    if (!CheckPwd(body, serverPassword))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("uuid").GetString() ?? "";
    var signature = body.GetProperty("signature").GetString() ?? "";
    var challenge = body.GetProperty("challenge").GetString() ?? "";
    var pubKey = GetPublicKey(db, uuid);
    if (pubKey == null) return Results.Json(new { error = "unknown machine" }, statusCode: 403);
    if (!VerifySignature(pubKey, challenge, signature))
        return Results.Json(new { error = "invalid signature" }, statusCode: 403);

    var machine = await db.Machines.FindAsync(uuid);
    if (machine != null) machine.LastSeen = DateTime.Now.ToString("O");

    // 获取该设备所有活跃的签到任务
    var tasks = await db.SignInTasks
        .Where(s => s.MachineUuid == uuid && s.Status == "active")
        .ToListAsync();

    var results = new List<object>();
    foreach (var t in tasks)
    {
        var records = JsonSerializer.Deserialize<List<SignInRecord>>(t.SignInRecords) ?? new();
        results.Add(new
        {
            short_code = t.ShortCode,
            task_id = $"signin_{t.ShortCode}",
            classroom = t.Classroom,
            subject = t.Subject,
            student_list = JsonSerializer.Deserialize<List<string>>(t.StudentList) ?? new(),
            records
        });
    }

    await db.SaveChangesAsync();
    return Results.Json(new { tasks = results });
});

// ---- 签到记录数据模型 ----
/// <summary>
/// 单条签到记录（学生姓名 + 签到时间 + 设备标识）
/// </summary>
var signinRecordType = new { name = "", time = "", device = "" };

// ===== 学生签到页面 =====

/// <summary>
/// GET /s/{shortCode} - 学生签到页面（展示表单）
/// </summary>
app.MapGet("/s/{shortCode}", async (string shortCode, AppDbContext db, HttpContext ctx) =>
{
    var task = await db.SignInTasks.FirstOrDefaultAsync(s => s.ShortCode == shortCode);
    if (task == null)
        return Results.Content(RenderSignInPage(null, "签到任务不存在或已过期"), "text/html;charset=utf-8");

    if (task.Status != "active")
        return Results.Content(RenderSignInPage(null, "该签到任务已关闭"), "text/html;charset=utf-8");

    // 检查该设备是否已签到（通过 Cookie）
    var deviceCookie = $"si_dev_{shortCode}";
    if (ctx.Request.Cookies.ContainsKey(deviceCookie))
        return Results.Content(RenderSignInPage(task, null, "您已签到成功，无需重复签到"), "text/html;charset=utf-8");

    return Results.Content(RenderSignInPage(task), "text/html;charset=utf-8");
});

/// <summary>
/// POST /s/{shortCode} - 学生提交签到表单
/// </summary>
app.MapPost("/s/{shortCode}", async (string shortCode, AppDbContext db, HttpContext ctx) =>
{
    var task = await db.SignInTasks.FirstOrDefaultAsync(s => s.ShortCode == shortCode);
    if (task == null)
        return Results.Content(RenderSignInPage(null, "签到任务不存在或已过期"), "text/html;charset=utf-8");

    if (task.Status != "active")
        return Results.Content(RenderSignInPage(null, "该签到任务已关闭"), "text/html;charset=utf-8");

    var deviceCookie = $"si_dev_{shortCode}";
    if (ctx.Request.Cookies.ContainsKey(deviceCookie))
        return Results.Content(RenderSignInPage(task, null, "您已签到成功，无需重复签到"), "text/html;charset=utf-8");

    var form = await ctx.Request.ReadFormAsync();
    var name = form["name"].ToString().Trim();
    var classroom = form["classroom"].ToString().Trim();
    var password = form["password"].ToString().Trim();

    // 验证必填字段
    var errors = new List<string>();
    if (string.IsNullOrEmpty(name)) errors.Add("请输入姓名");
    if (string.IsNullOrEmpty(classroom)) errors.Add("请输入教室");
    if (string.IsNullOrEmpty(password)) errors.Add("请输入签到密码");
    if (errors.Count > 0)
        return Results.Content(RenderSignInPage(task, string.Join("；", errors)), "text/html;charset=utf-8");

    // 验证密码
    if (password != task.Password)
        return Results.Content(RenderSignInPage(task, "签到密码错误"), "text/html;charset=utf-8");

    // 检查是否在学生名单中
    var studentList = JsonSerializer.Deserialize<List<string>>(task.StudentList) ?? new();
    if (studentList.Count > 0 && !studentList.Contains(name, StringComparer.OrdinalIgnoreCase))
        return Results.Content(RenderSignInPage(task, "您不在该签到任务的学生名单中"), "text/html;charset=utf-8");

    // 检查是否已经签到过（按姓名查重）
    var records = JsonSerializer.Deserialize<List<SignInRecord>>(task.SignInRecords) ?? new();
    if (records.Any(r => r.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
        return Results.Content(RenderSignInPage(task, null, "该姓名已签到，请勿重复签到"), "text/html;charset=utf-8");

    // 记录签到
    var nowStr = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    var deviceFingerprint = ctx.Connection.RemoteIpAddress?.ToString() ?? "unknown";
    records.Add(new SignInRecord
    {
        Name = name,
        Time = nowStr,
        Device = deviceFingerprint
    });
    task.SignInRecords = JsonSerializer.Serialize(records);

    // 同步更新 attendance 表中的数据
    var taskId = $"signin_{shortCode}";
    var latest = await db.AttendanceRecords
        .Where(a => a.MachineUuid == task.MachineUuid && a.TaskId == taskId)
        .OrderByDescending(a => a.UpdatedAt)
        .FirstOrDefaultAsync();

    var attendanceData = latest != null
        ? JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latest.Data) ?? new()
        : new();

    if (!attendanceData.ContainsKey(name))
        attendanceData[name] = new StudentAttendance { Name = name };
    attendanceData[name].FirstTime = nowStr;
    attendanceData[name].Count++;
    attendanceData[name].History.Add(nowStr);

    db.AttendanceRecords.Add(new AttendanceEntity
    {
        MachineUuid = task.MachineUuid,
        TaskId = taskId,
        Data = JsonSerializer.Serialize(attendanceData),
        UpdatedAt = DateTime.Now.ToString("O")
    });

    await db.SaveChangesAsync();

    // 设置设备 Cookie（30 天有效，防止同一设备重复签到）
    ctx.Response.Cookies.Append(deviceCookie, "1", new CookieOptions
    {
        HttpOnly = true,
        SameSite = SameSiteMode.Strict,
        MaxAge = TimeSpan.FromDays(30),
        Path = $"/s/{shortCode}"
    });

    return Results.Content(RenderSignInPage(task, null, $"签到成功！{name} 于 {nowStr} 完成签到"), "text/html;charset=utf-8");
});

/// <summary>
/// 渲染学生签到页面 HTML
/// </summary>
string RenderSignInPage(SignInTaskEntity? task, string? errorMsg = null, string? successMsg = null)
{
    if (task == null)
    {
        return $@"<!DOCTYPE html><html lang=""zh""><head><meta charset=""utf-8""><meta name=""viewport"" content=""width=device-width,initial-scale=1""><title>签到 - {HttpUtility.HtmlEncode(serverName)}</title>
<style>*{{margin:0;padding:0;box-sizing:border-box}}body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f0f4f8;min-height:100vh;display:flex;align-items:center;justify-content:center}}div{{background:#fff;border-radius:16px;padding:40px;box-shadow:0 4px 20px rgba(0,0,0,0.1);max-width:400px;width:90%;text-align:center}}h2{{color:#333;margin-bottom:16px}}p{{color:#666;font-size:14px}}</style></head>
<body><div><h2>😕 签到不可用</h2><p>{(successMsg != null ? HttpUtility.HtmlEncode(successMsg) : (errorMsg != null ? HttpUtility.HtmlEncode(errorMsg) : "签到任务不存在或已过期"))}</p></div></body></html>";
    }

    var hasError = !string.IsNullOrEmpty(errorMsg);
    var hasSuccess = !string.IsNullOrEmpty(successMsg);
    var errorDisplay = hasError ? "block" : "none";
    var successDisplay = hasSuccess ? "block" : "none";
    var formDisplay = (!hasError && !hasSuccess) ? "block" : "none";

    return $@"<!DOCTYPE html><html lang=""zh""><head><meta charset=""utf-8""><meta name=""viewport"" content=""width=device-width,initial-scale=1""><title>签到 - {HttpUtility.HtmlEncode(task.Subject)} - {HttpUtility.HtmlEncode(task.Classroom)}</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f0f4f8;min-height:100vh;display:flex;align-items:center;justify-content:center}}
.card{{background:#fff;border-radius:16px;padding:32px 28px;box-shadow:0 4px 24px rgba(0,0,0,0.08);max-width:420px;width:90%}}
h2{{color:#333;font-size:22px;text-align:center;margin-bottom:4px}}
h3{{color:#888;font-size:14px;text-align:center;font-weight:normal;margin-bottom:24px}}
label{{display:block;font-size:13px;color:#555;margin-bottom:6px;font-weight:500}}
input{{width:100%;padding:10px 14px;border:1.5px solid #e0e0e0;border-radius:8px;font-size:15px;outline:none;transition:border .2s;margin-bottom:16px}}
input:focus{{border-color:#4285f4}}
.btn{{width:100%;padding:12px;background:#4285f4;color:#fff;border:none;border-radius:8px;font-size:16px;font-weight:600;cursor:pointer;transition:background .2s}}
.btn:hover{{background:#3367d6}}
.alert{{padding:12px 16px;border-radius:8px;font-size:14px;margin-bottom:16px;display:none}}
.alert-error{{background:#fce8e6;color:#c5221f;display:{errorDisplay}}}
.alert-success{{background:#e6f4ea;color:#137333;display:{successDisplay}}}
.form-area{{display:{formDisplay}}}
.footer{{text-align:center;margin-top:16px;font-size:12px;color:#aaa}}
</style></head>
<body>
<div class=""card"">
<h2>{HttpUtility.HtmlEncode(task.Subject)} 课堂签到</h2>
<h3>教室：{HttpUtility.HtmlEncode(task.Classroom)}</h3>
<div class=""alert alert-error"">{HttpUtility.HtmlEncode(errorMsg ?? "")}</div>
<div class=""alert alert-success"">{HttpUtility.HtmlEncode(successMsg ?? "")}</div>
<div class=""form-area"">
<form method=""post"" action=""/s/{HttpUtility.HtmlEncode(task.ShortCode)}"">
<label>姓名</label>
<input type=""text"" name=""name"" placeholder=""请输入你的姓名"" required autofocus>
<label>教室</label>
<input type=""text"" name=""classroom"" placeholder=""请输入教室名称"" required>
<label>签到密码</label>
<input type=""password"" name=""password"" placeholder=""请输入教师提供的签到密码"" required>
<button type=""submit"" class=""btn"">确认签到</button>
</form>
</div>
<div class=""footer"">AgoraIn 签到系统</div>
</div>
</body></html>";
}

// ===== 认证页面 =====

/// <summary>
/// GET /login - 显示登录页面
/// </summary>
app.MapGet("/login", () => Results.Content(RenderLoginPage(), "text/html;charset=utf-8"));

/// <summary>
/// POST /login - 提交登录表单，从 Users 表验证用户名和 SHA256 密码哈希
/// </summary>
app.MapPost("/login", async (HttpContext ctx, AppDbContext db) =>
{
    var form = await ctx.Request.ReadFormAsync();
    var username = form["username"].ToString();
    var password = form["password"].ToString();

    // 从数据库查询用户并验证密码
    var passwordHash = Sha256(password);
    var user = await db.Users.FirstOrDefaultAsync(u => u.Username == username && u.PasswordHash == passwordHash && u.IsActive);

    if (user != null)
    {
        var token = MakeToken(username, user.Role);
        ctx.Response.Cookies.Append(CookieName, token, new CookieOptions
        {
            HttpOnly = true,
            SameSite = SameSiteMode.Strict,
            MaxAge = TimeSpan.FromDays(7),
            Path = "/"
        });
        return Results.Redirect("/");
    }

    return Results.Content(RenderLoginPage("用户名或密码错误"), "text/html;charset=utf-8");
});

/// <summary>
/// GET /logout - 清除会话 Cookie 并重定向到登录页
/// </summary>
app.MapGet("/logout", (HttpContext ctx) =>
{
    ctx.Response.Cookies.Delete(CookieName);
    return Results.Redirect("/login");
});

// ===== Web 管理面板页面 =====

/// <summary>
/// GET / - 首页：显示所有已注册设备的列表（文件夹式卡片布局）
/// </summary>
app.MapGet("/", async (AppDbContext db, HttpContext ctx) =>
{
    var machines = await db.Machines.ToListAsync();
    var attendances = await db.AttendanceRecords.ToListAsync();
    var now = DateTime.Now;
    int onlineCount = 0, totalCount = machines.Count;

    var rows = new StringBuilder();
    foreach (var m in machines)
    {
        var last = m.LastSeen != null ? DateTime.Parse(m.LastSeen) : (DateTime?)null;
        var online = last != null && (now - last.Value).TotalSeconds < 300;
        if (online) onlineCount++;

        var taskCount = attendances.Where(a => a.MachineUuid == m.Uuid).Select(a => a.TaskId).Distinct().Count();
        var badgeCls = online ? "badge-online" : "badge-offline";
        var badgeText = online ? "在线" : "离线";
        var lastSeenStr = last?.ToString("yyyy-MM-dd HH:mm:ss") ?? "从未连接";

        // 文件夹样式显示
        rows.Append($@"<tr class=""machine-row"" onclick=""window.location.href='/machine/{m.Uuid}'"">
<td><div class=""folder-icon""><span class=""icon"">📁</span><span class=""name"">{HttpUtility.HtmlEncode(m.Name)}</span></div></td>
<td><span class=""badge {badgeCls}"">{badgeText}</span></td>
<td>{taskCount} 个任务</td>
<td>{lastSeenStr}</td>
<td onclick=""event.stopPropagation();event.preventDefault()"">
<div style=""display:flex;gap:6px;"">
<button class=""btn btn-sm"" onclick=""event.stopPropagation();window.location.href='/machine/{m.Uuid}'"">查看</button>
<button class=""btn btn-sm btn-danger"" onclick=""event.stopPropagation();openDeleteMachineModal('{m.Uuid}','{HttpUtility.HtmlEncode(m.Name)}')"">删除</button>
</div>
</td>
</tr>");
    }

    var content = $@"<div class=""page-header""><h2>设备总览</h2></div>
<div class=""stats-row"">
    <div class=""stat-card""><div class=""label"">设备总数</div><div class=""value"">{totalCount}</div></div>
    <div class=""stat-card online""><div class=""label"">在线设备</div><div class=""value"">{onlineCount}</div></div>
    <div class=""stat-card offline""><div class=""label"">离线设备</div><div class=""value"">{totalCount - onlineCount}</div></div>
</div>
<div class=""card"">
    <h3>已注册设备</h3>
    {(totalCount > 0 ? $"<table><thead><tr><th>设备名称</th><th>状态</th><th>任务数</th><th>最后在线</th><th>操作</th></tr></thead><tbody>{rows}</tbody></table>" : "<p style='color:var(--text-secondary);padding:20px 0;'>暂无已注册设备</p>")}
</div>
<style>
.machine-row {{ cursor:pointer; transition:background 0.15s; }}
.machine-row:hover {{ background:#f0f4f8; }}
.folder-icon {{ display:flex;align-items:center;gap:8px; }}
.folder-icon .icon {{ font-size:18px; }}
.folder-icon .name {{ font-weight:500; }}
</style>";
    return Results.Content(RenderPage(content, "home", ctx), "text/html;charset=utf-8");
});

/// <summary>
/// GET /machine/{uuid} - 设备详情页：显示该设备的所有打卡任务（卡片式布局）
/// </summary>
app.MapGet("/machine/{uuid}", async (string uuid, AppDbContext db, HttpContext ctx) =>
{
    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Content(RenderPage("<div class='card'><h3>设备不存在</h3><p>该 UUID 对应的设备不存在。</p></div>", ctx: ctx), "text/html;charset=utf-8");

    var config = JsonSerializer.Deserialize<ClientConfig>(machine.Config) ?? new ClientConfig();
    var taskRecords = await db.AttendanceRecords
        .Where(a => a.MachineUuid == uuid)
        .ToListAsync();

    var last = machine.LastSeen != null ? DateTime.Parse(machine.LastSeen) : (DateTime?)null;
    var online = last != null && (DateTime.Now - last.Value).TotalSeconds < 300;
    var badgeCls = online ? "badge-online" : "badge-offline";
    var badgeText = online ? "在线" : "离线";

    // 任务卡片列表
    var taskCards = new StringBuilder();
    foreach (var tg in taskRecords.GroupBy(a => a.TaskId))
    {
        var taskId = tg.Key;
        var latestRecord = tg.OrderByDescending(a => a.UpdatedAt).FirstOrDefault();
        var taskTime = latestRecord?.UpdatedAt ?? "未知";
        var taskData = latestRecord != null ? JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latestRecord.Data) ?? new() : new();
        var totalStudents = taskData.Count;
        var punchedCount = taskData.Values.Count(v => v.FirstTime != null);
        var taskDisplayName = string.IsNullOrEmpty(taskId) || taskId == "default" ? "默认任务" : taskId;

        taskCards.Append($@"<div class=""task-card"" onclick=""window.location.href='/machine/{uuid}/task/{HttpUtility.UrlEncode(taskId)}'"">
<div class=""task-header"">
<span class=""task-icon"">📋</span>
<span class=""task-name"">{HttpUtility.HtmlEncode(taskDisplayName)}</span>
</div>
<div class=""task-stats"">
<div class=""stat""><span class=""label"">总人数</span><span class=""value"">{totalStudents}</span></div>
<div class=""stat""><span class=""label"">已打卡</span><span class=""value"">{punchedCount}</span></div>
</div>
<div class=""task-time"">最后更新: {taskTime}</div>
</div>");
    }

    if (taskRecords.Count == 0)
    {
        taskCards.Append("<p style='color:var(--text-secondary);padding:20px 0;'>该设备暂无打卡任务</p>");
    }

    var escapedConfig = HttpUtility.HtmlAttributeEncode(machine.Config);

    var content = $@"<div class=""breadcrumb""><a href=""/"">设备总览</a> / {HttpUtility.HtmlEncode(machine.Name)}</div>
<div class=""page-header"">
<h2>{HttpUtility.HtmlEncode(machine.Name)}</h2>
<div style=""display:flex;gap:8px;"">
<span class=""badge {badgeCls}"">{badgeText}</span>
<button class=""btn btn-sm"" onclick=""openEditConfigModal('{uuid}','{escapedConfig}')"">编辑配置</button>
<button class=""btn btn-sm btn-danger"" onclick=""openDeleteMachineModal('{uuid}','{HttpUtility.HtmlEncode(machine.Name)}')"">删除设备</button>
</div>
</div>

<div class=""card"">
    <h3>任务列表</h3>
    <p style=""color:var(--text-secondary);font-size:13px;margin-bottom:16px;"">点击任务查看详细打卡数据</p>
    <div class=""task-grid"">{taskCards}</div>
</div>

<style>
.task-grid {{ display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:16px; }}
.task-card {{ background:#fff;border:1.5px solid var(--border);border-radius:12px;padding:16px;cursor:pointer;transition:all 0.2s; }}
.task-card:hover {{ transform:translateY(-3px);box-shadow:0 8px 20px rgba(0,0,0,0.1);border-color:var(--primary); }}
.task-header {{ display:flex;align-items:center;gap:8px;margin-bottom:12px; }}
.task-icon {{ font-size:20px; }}
.task-name {{ font-weight:600;font-size:15px;color:var(--text); }}
.task-stats {{ display:flex;gap:20px;margin-bottom:12px; }}
.task-stats .stat {{ display:flex;flex-direction:column; }}
.task-stats .label {{ font-size:12px;color:var(--text-secondary); }}
.task-stats .value {{ font-size:18px;font-weight:600;color:var(--text); }}
.task-time {{ font-size:12px;color:var(--text-secondary); }}
</style>";
    return Results.Content(RenderPage(content, "home", ctx), "text/html;charset=utf-8");
});

/// <summary>
/// GET /machine/{uuid}/task/{taskId} - 任务详情页：显示打卡排名和状态网格，支持 Web 打卡
/// </summary>
app.MapGet("/machine/{uuid}/task/{taskId}", async (string uuid, string taskId, AppDbContext db, HttpContext ctx) =>
{
    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Content(RenderPage("<div class='card'><h3>设备不存在</h3></div>", ctx: ctx), "text/html;charset=utf-8");

    var decodedTaskId = HttpUtility.UrlDecode(taskId);
    var latest = await db.AttendanceRecords
        .Where(a => a.MachineUuid == uuid && a.TaskId == decodedTaskId)
        .OrderByDescending(a => a.UpdatedAt)
        .FirstOrDefaultAsync();

    var data = latest != null ? JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latest.Data) ?? new() : new();
    var updateTime = latest?.UpdatedAt ?? "从未同步";

    var punched = data.Values.Where(d => d.FirstTime != null).OrderBy(d => d.FirstTime).ToList();
    int totalStudents = data.Count, punchedCount = punched.Count;

    var rankRows = new StringBuilder();
    int i = 1;
    foreach (var p in punched)
    {
        var timeStr = p.FirstTime != null && p.FirstTime.Length >= 16 ? p.FirstTime[11..16] : (p.FirstTime ?? "--:--");
        rankRows.Append($"<tr><td><strong>{i++}</strong></td><td>{HttpUtility.HtmlEncode(p.Name)}</td><td>{timeStr}</td></tr>");
    }
    var rankTable = rankRows.Length > 0 ? $"<table><thead><tr><th>排名</th><th>姓名</th><th>时间</th></tr></thead><tbody>{rankRows}</tbody></table>" : "<p style='color:var(--text-secondary);'>暂无打卡记录</p>";

    var gridItems = new StringBuilder();
    foreach (var (name, sa) in data)
    {
        var cls = sa.FirstTime != null ? "punched" : "";
        var st = sa.FirstTime != null ? "true" : "false";
        gridItems.Append($"<div class=\"grid-item {cls}\" onclick=\"openPunchModal('{uuid}','{HttpUtility.UrlEncode(decodedTaskId)}','{HttpUtility.HtmlEncode(name)}',{st})\">{HttpUtility.HtmlEncode(name)}</div>");
    }

    var taskDisplayName = string.IsNullOrEmpty(decodedTaskId) || decodedTaskId == "default" ? "默认任务" : decodedTaskId;

    var content = $@"<div class=""breadcrumb""><a href=""/"">设备总览</a> / <a href=""/machine/{uuid}"">{HttpUtility.HtmlEncode(machine.Name)}</a> / {HttpUtility.HtmlEncode(taskDisplayName)}</div>
<div class=""page-header""><h2>{HttpUtility.HtmlEncode(taskDisplayName)}</h2><div style=""display:flex;gap:8px;"">
    <button class=""btn btn-sm btn-danger"" onclick=""openClearDataModal('{uuid}','{HttpUtility.UrlEncode(decodedTaskId)}')"">清除数据</button>
</div></div>

<div class=""stats-row"">
    <div class=""stat-card""><div class=""label"">总人数</div><div class=""value"">{totalStudents}</div></div>
    <div class=""stat-card online""><div class=""label"">已打卡</div><div class=""value"">{punchedCount}</div></div>
    <div class=""stat-card""><div class=""label"">未打卡</div><div class=""value"">{totalStudents - punchedCount}</div></div>
</div>

<div style=""display:grid;grid-template-columns:1fr 1fr;gap:20px;"">
    <div class=""card""><h3>打卡排名</h3>{rankTable}</div>
    <div class=""card""><h3>打卡状态</h3><div class=""student-grid"">{gridItems}</div></div>
</div>";
    return Results.Content(RenderPage(content, "home", ctx), "text/html;charset=utf-8");
});

// ===== 用户管理页面 =====

/// <summary>
/// GET /users - 用户管理页面（仅管理员可见）
/// </summary>
app.MapGet("/users", async (AppDbContext db, HttpContext ctx) =>
{
    if (!IsAdmin(ctx))
        return Results.Content(RenderPage("<div class='card'><h3>权限不足</h3><p>此页面仅管理员可访问。</p></div>", "home", ctx), "text/html;charset=utf-8");

    var users = await db.Users.OrderBy(u => u.Id).ToListAsync();
    var rows = new StringBuilder();

    foreach (var u in users)
    {
        var roleLabel = u.Role switch
        {
            "admin" => "<span class='badge badge-online'>管理员</span>",
            "operator" => "<span class='badge' style='background:#e8f0fe;color:#1967d2'>操作员</span>",
            _ => "<span class='badge badge-offline' style='background:#f3e8fd;color:#7c3aed'>查看者</span>"
        };
        var statusBadge = u.IsActive
            ? "<span class='badge badge-online'>启用</span>"
            : "<span class='badge badge-offline'>禁用</span>";

        rows.Append($@"<tr>
<td>{u.Id}</td>
<td>{HttpUtility.HtmlEncode(u.Username)}</td>
<td>{HttpUtility.HtmlEncode(u.DisplayName)}</td>
<td>{roleLabel}</td>
<td>{statusBadge}</td>
<td>{u.CreatedAt}</td>
<td>
<div style=""display:flex;gap:6px;"">
<button class=""btn btn-sm"" onclick=""openEditUserModal({u.Id},'{HttpUtility.HtmlEncode(u.Username)}','{HttpUtility.HtmlEncode(u.DisplayName)}','{u.Role}',{u.IsActive.ToString().ToLower()})"">编辑</button>
<button class=""btn btn-sm btn-danger"" onclick=""deleteUser({u.Id},'{HttpUtility.HtmlEncode(u.Username)}')"">删除</button>
</div>
</td>
</tr>");
    }

    var content = $@"<div class=""page-header""><h2>用户管理</h2>
<button class=""btn"" onclick=""openCreateUserModal()"">添加用户</button>
</div>
<div class=""card"">
<h3>用户列表</h3>
{(users.Count > 0 ? $"<table><thead><tr><th>ID</th><th>用户名</th><th>显示名称</th><th>角色</th><th>状态</th><th>创建时间</th><th>操作</th></tr></thead><tbody>{rows}</tbody></table>" : "<p style='color:var(--text-secondary);padding:20px 0;'>暂无用户</p>")}
</div>
<script>
function openCreateUserModal() {{
    var html = '<h3>创建用户</h3>' +
        '<form id=""createUserForm"">' +
        '<div class=""form-group""><label>用户名</label><input name=""username"" placeholder=""请输入用户名"" required></div>' +
        '<div class=""form-group""><label>密码</label><input type=""password"" name=""password"" placeholder=""请输入密码"" required></div>' +
        '<div class=""form-group""><label>显示名称</label><input name=""display_name"" placeholder=""请输入显示名称""></div>' +
        '<div class=""form-group""><label>角色</label><select name=""role""><option value=""viewer"">查看者</option><option value=""operator"">操作员</option><option value=""admin"">管理员</option></select></div>' +
        '<div class=""form-actions""><button type=""submit"" class=""btn"">创建</button><button type=""button"" class=""btn btn-ghost"" onclick=""closeModal(\'modal\')"">取消</button></div></form>';
    document.getElementById('modal-body').innerHTML = html;
    document.getElementById('modal').style.display = 'block';
    document.getElementById('createUserForm').onsubmit = function(e) {{
        e.preventDefault();
        var fd = new FormData(e.target);
        var data = {{ username: fd.get('username'), password: fd.get('password'), display_name: fd.get('display_name'), role: fd.get('role') }};
        fetch('/api/users', {{ method: 'POST', headers: {{'Content-Type':'application/json'}}, body: JSON.stringify(data) }})
            .then(r => r.json())
            .then(d => {{
                if (d.error) {{ showToast('创建失败: ' + d.error, 'error'); }}
                else {{ showToast('用户创建成功', 'success'); setTimeout(() => location.reload(), 800); }}
            }})
            .catch(() => showToast('网络请求失败', 'error'));
    }};
}}

function openEditUserModal(id, username, displayName, role, isActive) {{
    var activeChecked = isActive ? 'checked' : '';
    var html = '<h3>编辑用户 - ' + username + '</h3>' +
        '<form id=""editUserForm"">' +
        '<div class=""form-group""><label>显示名称</label><input name=""display_name"" value=""' + displayName + '""></div>' +
        '<div class=""form-group""><label>角色</label><select name=""role""><option value=""viewer""' + (role==='viewer'?' selected':'') + '>查看者</option><option value=""operator""' + (role==='operator'?' selected':'') + '>操作员</option><option value=""admin""' + (role==='admin'?' selected':'') + '>管理员</option></select></div>' +
        '<div class=""form-group""><label><input type=""checkbox"" name=""is_active"" ' + activeChecked + '> 启用账户</label></div>' +
        '<div class=""form-actions""><button type=""submit"" class=""btn"">保存</button><button type=""button"" class=""btn btn-ghost"" onclick=""closeModal(\'modal\')"">取消</button></div></form>';
    document.getElementById('modal-body').innerHTML = html;
    document.getElementById('modal').style.display = 'block';
    document.getElementById('editUserForm').onsubmit = function(e) {{
        e.preventDefault();
        var fd = new FormData(e.target);
        var data = {{ display_name: fd.get('display_name'), role: fd.get('role'), is_active: fd.get('is_active') === 'on' }};
        fetch('/api/users/' + id, {{ method: 'PUT', headers: {{'Content-Type':'application/json'}}, body: JSON.stringify(data) }})
            .then(r => r.json())
            .then(d => {{
                if (d.error) {{ showToast('更新失败: ' + d.error, 'error'); }}
                else {{ showToast('用户更新成功', 'success'); setTimeout(() => location.reload(), 800); }}
            }})
            .catch(() => showToast('网络请求失败', 'error'));
    }};
}}

function deleteUser(id, username) {{
    if (!confirm('确定要删除用户 ""' + username + '"" 吗？此操作不可撤销！')) return;
    fetch('/api/users/' + id, {{ method: 'DELETE' }})
        .then(r => r.json())
        .then(d => {{
            if (d.error) {{ showToast('删除失败: ' + d.error, 'error'); }}
            else {{ showToast('用户已删除', 'success'); setTimeout(() => location.reload(), 800); }}
        }})
        .catch(() => showToast('网络请求失败', 'error'));
}}
</script>";
    return Results.Content(RenderPage(content, "users", ctx), "text/html;charset=utf-8");
});

/// <summary>
/// GET /profile - 个人设置页面（修改密码）
/// </summary>
app.MapGet("/profile", async (HttpContext ctx, AppDbContext db) =>
{
    var username = GetUsername(ctx);
    if (username == null)
        return Results.Content(RenderPage("<div class='card'><h3>未登录</h3></div>", "profile", ctx), "text/html;charset=utf-8");

    var user = await db.Users.FirstOrDefaultAsync(u => u.Username == username);
    if (user == null)
        return Results.Content(RenderPage("<div class='card'><h3>用户不存在</h3></div>", "profile", ctx), "text/html;charset=utf-8");

    var roleLabel = user.Role switch
    {
        "admin" => "管理员",
        "operator" => "操作员",
        _ => "查看者"
    };

    var content = $@"<div class=""page-header""><h2>个人设置</h2></div>
<div class=""card"" style=""max-width:500px;"">
<h3>账户信息</h3>
<ul class=""config-list"">
<li><span class=""label"">用户名</span><span class=""value"">{HttpUtility.HtmlEncode(user.Username)}</span></li>
<li><span class=""label"">显示名称</span><span class=""value"">{HttpUtility.HtmlEncode(user.DisplayName)}</span></li>
<li><span class=""label"">角色</span><span class=""value"">{roleLabel}</span></li>
<li><span class=""label"">创建时间</span><span class=""value"">{user.CreatedAt}</span></li>
</ul>
</div>
<div class=""card"" style=""max-width:500px;"">
<h3>修改密码</h3>
<form id=""changePwdForm"">
<div class=""form-group""><label>旧密码</label><input type=""password"" id=""oldPwd"" placeholder=""请输入旧密码"" required></div>
<div class=""form-group""><label>新密码</label><input type=""password"" id=""newPwd"" placeholder=""请输入新密码"" required></div>
<div class=""form-group""><label>确认新密码</label><input type=""password"" id=""confirmPwd"" placeholder=""请再次输入新密码"" required></div>
<div class=""form-actions""><button type=""submit"" class=""btn"">修改密码</button></div>
</form>
</div>
<script>
document.getElementById('changePwdForm').onsubmit = function(e) {{
    e.preventDefault();
    var oldPwd = document.getElementById('oldPwd').value;
    var newPwd = document.getElementById('newPwd').value;
    var confirmPwd = document.getElementById('confirmPwd').value;
    if (!oldPwd || !newPwd || !confirmPwd) {{ showToast('请填写所有字段', 'error'); return; }}
    if (newPwd !== confirmPwd) {{ showToast('两次输入的新密码不一致', 'error'); return; }}
    fetch('/api/users/change-password', {{
        method: 'POST',
        headers: {{'Content-Type':'application/json'}},
        body: JSON.stringify({{ old_password: oldPwd, new_password: newPwd }})
    }})
    .then(r => r.json())
    .then(d => {{
        if (d.error) {{ showToast('修改失败: ' + d.error, 'error'); }}
        else {{ showToast('密码修改成功，请重新登录', 'success'); setTimeout(() => location.href = '/logout', 1500); }}
    }})
    .catch(() => showToast('网络请求失败', 'error'));
}};
</script>";
    return Results.Content(RenderPage(content, "profile", ctx), "text/html;charset=utf-8");
});

// =============================================================================
// 移动端 API（Bearer Token 认证，供 Admin 和学生端 App 使用）
// =============================================================================

// ---- 认证 API ----

/// <summary>
/// POST /api/auth/login - 移动端登录，验证用户名密码，返回 Bearer Token
/// </summary>
app.MapPost("/api/auth/login", async (AppDbContext db, JsonElement body) =>
{
    var username = body.TryGetProperty("username", out var un) ? un.GetString()?.Trim() ?? "" : "";
    var password = body.TryGetProperty("password", out var pw) ? pw.GetString() ?? "" : "";

    if (string.IsNullOrEmpty(username) || string.IsNullOrEmpty(password))
        return Results.Json(new { error = "用户名和密码不能为空" }, statusCode: 400);

    var passwordHash = Sha256(password);
    var user = await db.Users.FirstOrDefaultAsync(u => u.Username == username && u.PasswordHash == passwordHash);

    if (user == null)
        return Results.Json(new { error = "用户名或密码错误" }, statusCode: 401);

    if (!user.IsActive)
        return Results.Json(new { error = "该账户已被禁用，请联系管理员" }, statusCode: 403);

    var token = MakeLoginToken(username, user.Role);

    return Results.Json(new
    {
        token,
        user = new
        {
            id = user.Id,
            username = user.Username,
            role = user.Role,
            display_name = user.DisplayName
        }
    });
});

/// <summary>
/// POST /api/auth/verify - 验证 Bearer Token 是否有效，返回用户信息
/// </summary>
app.MapPost("/api/auth/verify", (HttpContext ctx) =>
{
    var (username, role, error) = ParseBearerToken(ctx);
    if (error != null)
        return Results.Json(new { valid = false, error }, statusCode: 401);

    return Results.Json(new { valid = true, username, role });
});

// ---- 二维码签到 API ----

/// <summary>
/// POST /api/qrcode/generate - 管理员创建签到任务，生成二维码数据（返回 shortCode）
/// 需要 admin 或 operator 角色
/// </summary>
app.MapPost("/api/qrcode/generate", async (AppDbContext db, HttpContext ctx) =>
{
    // Bearer Token 认证 + 权限检查
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (role != "admin" && role != "operator")
        return Results.Json(new { error = "权限不足，仅管理员或操作员可创建签到任务" }, statusCode: 403);

    string bodyStr;
    using (var reader = new StreamReader(ctx.Request.Body, Encoding.UTF8))
        bodyStr = await reader.ReadToEndAsync();

    JsonElement body;
    try { body = JsonDocument.Parse(bodyStr).RootElement; }
    catch { return Results.Json(new { error = "无效的 JSON 格式" }, statusCode: 400); }

    var classroom = body.TryGetProperty("classroom", out var cr) ? cr.GetString()?.Trim() ?? "" : "";
    var subject = body.TryGetProperty("subject", out var sj) ? sj.GetString()?.Trim() ?? "" : "";
    var signPassword = body.TryGetProperty("sign_password", out var sp) ? sp.GetString()?.Trim() ?? "" : "";
    var studentNamesStr = body.TryGetProperty("students", out var sn) ? sn.GetRawText() : "[]";
    var machineUuid = body.TryGetProperty("machine_uuid", out var mu) ? mu.GetString()?.Trim() ?? "" : "";

    if (string.IsNullOrEmpty(subject))
        return Results.Json(new { error = "科目名称不能为空" }, statusCode: 400);
    if (string.IsNullOrEmpty(signPassword))
        return Results.Json(new { error = "签到密码不能为空" }, statusCode: 400);

    // 解析学生名单
    List<string> studentNames;
    try { studentNames = JsonSerializer.Deserialize<List<string>>(studentNamesStr) ?? new(); }
    catch { return Results.Json(new { error = "学生名单格式错误" }, statusCode: 400); }

    // 如果没有指定设备 UUID，使用管理员用户名作为虚拟设备标识
    var actualUuid = string.IsNullOrEmpty(machineUuid)
        ? $"admin_{username}"
        : machineUuid;

    // 确保虚拟设备存在
    if (!await db.Machines.AnyAsync(m => m.Uuid == actualUuid))
    {
        db.Machines.Add(new MachineEntity
        {
            Uuid = actualUuid,
            Name = $"管理员 {username} 创建的签到",
            PublicKey = "mobile-admin",
            LastSeen = DateTime.Now.ToString("O")
        });
    }

    // 生成唯一短链码
    string shortCode;
    do { shortCode = GenerateShortCode(); }
    while (await db.SignInTasks.AnyAsync(s => s.ShortCode == shortCode));

    var task = new SignInTaskEntity
    {
        ShortCode = shortCode,
        MachineUuid = actualUuid,
        Password = signPassword,
        Classroom = classroom,
        Subject = subject,
        StudentList = studentNamesStr,
        SignInRecords = "[]",
        CreatedAt = DateTime.Now.ToString("O"),
        Status = "active"
    };
    db.SignInTasks.Add(task);

    // 同步创建 attendance 记录
    var taskId = $"signin_{shortCode}";
    var initialData = new Dictionary<string, StudentAttendance>();
    foreach (var name in studentNames)
        initialData[name] = new StudentAttendance { Name = name };

    db.AttendanceRecords.Add(new AttendanceEntity
    {
        MachineUuid = actualUuid,
        TaskId = taskId,
        Data = JsonSerializer.Serialize(initialData),
        UpdatedAt = DateTime.Now.ToString("O")
    });

    var machine = await db.Machines.FindAsync(actualUuid);
    if (machine != null) machine.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();

    return Results.Json(new
    {
        short_code = shortCode,
        task_id = taskId,
        qrcode_url = $"/s/{shortCode}",
        subject,
        classroom,
        student_count = studentNames.Count,
        created_at = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
    });
});

/// <summary>
/// POST /api/qrcode/checkin - 学生通过扫码签到（JSON API）
/// 任何拥有有效令牌的登录用户都可以使用（学生端登录即可）
/// </summary>
app.MapPost("/api/qrcode/checkin", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);

    string bodyStr;
    using (var reader = new StreamReader(ctx.Request.Body, Encoding.UTF8))
        bodyStr = await reader.ReadToEndAsync();

    JsonElement body;
    try { body = JsonDocument.Parse(bodyStr).RootElement; }
    catch { return Results.Json(new { error = "无效的 JSON 格式" }, statusCode: 400); }

    var shortCode = body.TryGetProperty("short_code", out var sc) ? sc.GetString()?.Trim() ?? "" : "";
    var studentName = body.TryGetProperty("student_name", out var sn) ? sn.GetString()?.Trim() ?? "" : "";
    var signPassword = body.TryGetProperty("sign_password", out var sp) ? sp.GetString()?.Trim() ?? "" : "";

    if (string.IsNullOrEmpty(shortCode))
        return Results.Json(new { error = "签到码不能为空" }, statusCode: 400);
    if (string.IsNullOrEmpty(studentName))
        return Results.Json(new { error = "姓名不能为空" }, statusCode: 400);
    if (string.IsNullOrEmpty(signPassword))
        return Results.Json(new { error = "签到密码不能为空" }, statusCode: 400);

    var task = await db.SignInTasks.FirstOrDefaultAsync(s => s.ShortCode == shortCode);
    if (task == null)
        return Results.Json(new { error = "签到任务不存在或已过期" }, statusCode: 404);

    if (task.Status != "active")
        return Results.Json(new { error = "该签到任务已关闭" }, statusCode: 400);

    // 验证密码
    if (signPassword != task.Password)
        return Results.Json(new { error = "签到密码错误" }, statusCode: 403);

    // 检查是否在学生名单中
    var studentList = JsonSerializer.Deserialize<List<string>>(task.StudentList) ?? new();
    if (studentList.Count > 0 && !studentList.Contains(studentName, StringComparer.OrdinalIgnoreCase))
        return Results.Json(new { error = "你不在该签到任务的学生名单中" }, statusCode: 403);

    // 检查是否已签到
    var records = JsonSerializer.Deserialize<List<SignInRecord>>(task.SignInRecords) ?? new();
    if (records.Any(r => r.Name.Equals(studentName, StringComparison.OrdinalIgnoreCase)))
        return Results.Json(new { error = "该姓名已签到，请勿重复签到" }, statusCode: 409);

    // 记录签到
    var nowStr = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    var deviceInfo = ctx.Connection.RemoteIpAddress?.ToString() ?? "mobile";
    records.Add(new SignInRecord
    {
        Name = studentName,
        Time = nowStr,
        Device = deviceInfo
    });
    task.SignInRecords = JsonSerializer.Serialize(records);

    // 同步更新 attendance 表
    var taskId = $"signin_{shortCode}";
    var latest = await db.AttendanceRecords
        .Where(a => a.MachineUuid == task.MachineUuid && a.TaskId == taskId)
        .OrderByDescending(a => a.UpdatedAt)
        .FirstOrDefaultAsync();

    var attendanceData = latest != null
        ? JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latest.Data) ?? new()
        : new();

    if (!attendanceData.ContainsKey(studentName))
        attendanceData[studentName] = new StudentAttendance { Name = studentName };
    attendanceData[studentName].FirstTime = nowStr;
    attendanceData[studentName].Count++;
    attendanceData[studentName].History.Add(nowStr);

    db.AttendanceRecords.Add(new AttendanceEntity
    {
        MachineUuid = task.MachineUuid,
        TaskId = taskId,
        Data = JsonSerializer.Serialize(attendanceData),
        UpdatedAt = DateTime.Now.ToString("O")
    });

    await db.SaveChangesAsync();

    return Results.Json(new
    {
        status = "ok",
        message = "签到成功",
        student_name = studentName,
        time = nowStr,
        subject = task.Subject,
        classroom = task.Classroom,
        rank = records.Count
    });
});

// ---- 管理员仪表盘 API ----

/// <summary>
/// GET /api/mobile/dashboard - 管理员仪表盘数据
/// 需要 admin 或 operator 角色
/// </summary>
app.MapGet("/api/mobile/dashboard", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (role != "admin" && role != "operator")
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var machines = await db.Machines.ToListAsync();
    var attendances = await db.AttendanceRecords.ToListAsync();
    var signInTasks = await db.SignInTasks.ToListAsync();
    var now = DateTime.Now;

    // 设备统计
    var onlineCount = machines.Count(m =>
    {
        var last = m.LastSeen != null ? DateTime.Parse(m.LastSeen) : (DateTime?)null;
        return last != null && (now - last.Value).TotalSeconds < 300;
    });

    // 今日签到统计
    var todayStr = now.ToString("yyyy-MM-dd");
    var todayCheckins = signInTasks
        .SelectMany(s => JsonSerializer.Deserialize<List<SignInRecord>>(s.SignInRecords) ?? new())
        .Count(r => r.Time.StartsWith(todayStr));

    // 活跃签到任务
    var activeSignInTasks = signInTasks.Where(s => s.Status == "active").Select(s => new
    {
        short_code = s.ShortCode,
        subject = s.Subject,
        classroom = s.Classroom,
        student_count = (JsonSerializer.Deserialize<List<string>>(s.StudentList) ?? new()).Count,
        signed_count = (JsonSerializer.Deserialize<List<SignInRecord>>(s.SignInRecords) ?? new()).Count,
        created_at = s.CreatedAt
    }).ToList();

    // 按设备汇总任务数
    var deviceTasks = machines.Select(m => new
    {
        uuid = m.Uuid,
        name = m.Name,
        task_count = attendances.Where(a => a.MachineUuid == m.Uuid).Select(a => a.TaskId).Distinct().Count(),
        last = m.LastSeen != null ? DateTime.Parse(m.LastSeen) : (DateTime?)null,
        online = m.LastSeen != null ? (now - DateTime.Parse(m.LastSeen)).TotalSeconds < 300 : false
    }).OrderByDescending(d => d.online).ToList();

    return Results.Json(new
    {
        summary = new
        {
            total_devices = machines.Count,
            online_devices = onlineCount,
            total_users = await db.Users.CountAsync(),
            today_checkins = todayCheckins,
            active_signin_tasks = activeSignInTasks.Count
        },
        devices = deviceTasks,
        active_signin_tasks = activeSignInTasks
    });
});

/// <summary>
/// GET /api/mobile/attendance - 查询打卡数据
/// 支持参数：machine_uuid, task_id（可选筛选）
/// 需要 admin 或 operator 角色
/// </summary>
app.MapGet("/api/mobile/attendance", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (role != "admin" && role != "operator")
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var machineUuid = ctx.Request.Query["machine_uuid"].FirstOrDefault();
    var taskId = ctx.Request.Query["task_id"].FirstOrDefault();

    var query = db.AttendanceRecords.AsQueryable();
    if (!string.IsNullOrEmpty(machineUuid))
        query = query.Where(a => a.MachineUuid == machineUuid);
    if (!string.IsNullOrEmpty(taskId))
        query = query.Where(a => a.TaskId == taskId);

    var records = await query
        .OrderByDescending(a => a.UpdatedAt)
        .Take(200) // 限制返回数量
        .ToListAsync();

    // 按 device+task 分组，返回每条最新记录
    var grouped = records
        .GroupBy(r => new { r.MachineUuid, r.TaskId })
        .Select(g =>
        {
            var latest = g.OrderByDescending(r => r.UpdatedAt).First();
            var machineName = db.Machines.Where(m => m.Uuid == latest.MachineUuid).Select(m => m.Name).FirstOrDefault() ?? "";
            var data = JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latest.Data) ?? new();
            var totalStudents = data.Count;
            var punchedCount = data.Values.Count(v => v.FirstTime != null);

            return new
            {
                machine_uuid = latest.MachineUuid,
                machine_name = machineName,
                task_id = latest.TaskId,
                total_students = totalStudents,
                punched_count = punchedCount,
                unpunched_count = totalStudents - punchedCount,
                attendance_rate = totalStudents > 0 ? Math.Round((double)punchedCount / totalStudents * 100, 1) : 0,
                last_updated = latest.UpdatedAt,
                students = data.Select(kv => new
                {
                    name = kv.Key,
                    checked_in = kv.Value.FirstTime != null,
                    first_time = kv.Value.FirstTime,
                    count = kv.Value.Count
                }).ToList()
            };
        }).ToList();

    return Results.Json(new { tasks = grouped });
});

/// <summary>
/// GET /api/mobile/tasks - 获取所有签到任务列表（管理员）
/// </summary>
app.MapGet("/api/mobile/tasks", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (role != "admin" && role != "operator")
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var tasks = await db.SignInTasks
        .OrderByDescending(s => s.CreatedAt)
        .Take(100)
        .Select(s => new
        {
            id = s.Id,
            short_code = s.ShortCode,
            subject = s.Subject,
            classroom = s.Classroom,
            status = s.Status,
            student_count = s.StudentList.Length > 2
                ? (JsonSerializer.Deserialize<List<string>>(s.StudentList) ?? new()).Count
                : 0,
            signed_count = s.SignInRecords.Length > 2
                ? (JsonSerializer.Deserialize<List<SignInRecord>>(s.SignInRecords) ?? new()).Count
                : 0,
            created_at = s.CreatedAt,
            machine_uuid = s.MachineUuid
        })
        .ToListAsync();

    return Results.Json(new { tasks });
});

/// <summary>
/// POST /api/mobile/tasks/{id}/close - 关闭签到任务
/// </summary>
app.MapPost("/api/mobile/tasks/{id}/close", async (int id, AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (role != "admin" && role != "operator")
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var task = await db.SignInTasks.FindAsync(id);
    if (task == null)
        return Results.Json(new { error = "任务不存在" }, statusCode: 404);

    task.Status = "closed";
    await db.SaveChangesAsync();

    return Results.Json(new { status = "ok", message = "任务已关闭" });
});

/// <summary>
/// DELETE /api/mobile/tasks/{id} - 删除签到任务及其打卡数据
/// </summary>
app.MapDelete("/api/mobile/tasks/{id}", async (int id, AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (role != "admin" && role != "operator")
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var task = await db.SignInTasks.FindAsync(id);
    if (task == null)
        return Results.Json(new { error = "任务不存在" }, statusCode: 404);

    var taskId = $"signin_{task.ShortCode}";
    db.AttendanceRecords.RemoveRange(
        db.AttendanceRecords.Where(a => a.MachineUuid == task.MachineUuid && a.TaskId == taskId));
    db.SignInTasks.Remove(task);
    await db.SaveChangesAsync();

    return Results.Json(new { status = "ok", message = "任务已删除" });
});

/// <summary>
/// GET /api/mobile/students - 获取学生签到历史（学生端查看自己的签到记录）
/// </summary>
app.MapGet("/api/mobile/students/history", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);

    // 查询该用户作为学生的签到记录（在所有签到任务中搜索）
    var signInTasks = await db.SignInTasks.ToListAsync();
    var history = new List<object>();

    foreach (var task in signInTasks)
    {
        var records = JsonSerializer.Deserialize<List<SignInRecord>>(task.SignInRecords) ?? new();
        var userRecord = records.FirstOrDefault(r => r.Name.Equals(username, StringComparison.OrdinalIgnoreCase));
        if (userRecord != null)
        {
            history.Add(new
            {
                subject = task.Subject,
                classroom = task.Classroom,
                short_code = task.ShortCode,
                time = userRecord.Time,
                status = task.Status
            });
        }
    }

    return Results.Json(new
    {
        student_name = username,
        total_checkins = history.Count,
        history = history.OrderByDescending(h => ((dynamic)h).time).ToList()
    });
});

app.Run();

/// <summary>
/// 单条签到记录（学生姓名 + 签到时间 + 设备标识）
/// </summary>
public class SignInRecord
{
    public string Name { get; set; } = "";
    public string Time { get; set; } = "";
    public string Device { get; set; } = "";
}
