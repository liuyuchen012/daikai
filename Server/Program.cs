using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using CheckIn.Server.Data;
using CheckIn.Shared.Models;
using Microsoft.EntityFrameworkCore;
using System.Web;

// =============================================================================
// SignWave 集控平台 - 服务器端入口
// 功能：设备注册与管理、打卡数据同步、Web 管理面板（含登录认证）
// =============================================================================

// ---- 加载服务器配置文件 config.json ----
var configPath = Path.Combine(AppContext.BaseDirectory, "config.json");
if (!File.Exists(configPath))
    configPath = Path.Combine(Directory.GetCurrentDirectory(), "config.json");

var configJson = File.Exists(configPath) ? JsonDocument.Parse(File.ReadAllText(configPath)).RootElement : default;
var cfgPort = configJson.TryGetProperty("Port", out var p) ? p.GetInt32() : 5250;
var cfgAdminUser = configJson.TryGetProperty("AdminUsername", out var au) ? au.GetString() ?? "admin" : "admin";
var cfgAdminPwd = configJson.TryGetProperty("AdminPassword", out var ap) ? ap.GetString() ?? "admin" : "admin";
var serverName = configJson.TryGetProperty("ServerName", out var sn) ? sn.GetString() ?? "SignWave 集控平台" : "SignWave 集控平台";
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
}

// ---- 会话认证（Session Auth） ----
// 使用 HMAC-SHA256 签名的 Cookie 实现无状态会话管理
var sessionSecret = Guid.NewGuid().ToString("N");
const string CookieName = "sw_session";

/// <summary>
/// 创建带签名的会话令牌，格式：username:timestamp:signature
/// </summary>
string MakeToken(string username)
{
    var ts = DateTimeOffset.UtcNow.ToUnixTimeSeconds().ToString();
    var payload = $"{username}:{ts}";
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
    if (parts.Length != 3) return false;
    var payload = $"{parts[0]}:{parts[1]}";
    using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(sessionSecret));
    var expected = Convert.ToHexString(hmac.ComputeHash(Encoding.UTF8.GetBytes(payload))).ToLower();
    return expected == parts[2];
}

/// <summary>
/// 检查当前请求是否已认证（通过 Cookie 中的会话令牌）
/// </summary>
bool IsAuthenticated(HttpContext ctx)
{
    if (!ctx.Request.Cookies.TryGetValue(CookieName, out var token)) return false;
    return ValidateToken(token);
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
string RenderPage(string content, string activeNav = "home")
{
    return templateContent
        .Replace("{TITLE}", HttpUtility.HtmlEncode(serverName))
        .Replace("{NAV_HOME}", activeNav == "home" ? "active" : "")
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
    // Allow: login page, logout, API endpoints, static files
    if (path.StartsWith("/api/") || path == "/login" || path == "/logout" || path.StartsWith("/static"))
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

// ===== 认证页面 =====

/// <summary>
/// GET /login - 显示登录页面
/// </summary>
app.MapGet("/login", () => Results.Content(RenderLoginPage(), "text/html;charset=utf-8"));

/// <summary>
/// POST /login - 提交登录表单，验证用户名密码后设置会话 Cookie
/// </summary>
app.MapPost("/login", async (HttpContext ctx) =>
{
    var form = await ctx.Request.ReadFormAsync();
    var username = form["username"].ToString();
    var password = form["password"].ToString();

    if (username == cfgAdminUser && password == cfgAdminPwd)
    {
        var token = MakeToken(username);
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
app.MapGet("/", async (AppDbContext db) =>
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
    return Results.Content(RenderPage(content, "home"), "text/html;charset=utf-8");
});

/// <summary>
/// GET /machine/{uuid} - 设备详情页：显示该设备的所有打卡任务（卡片式布局）
/// </summary>
app.MapGet("/machine/{uuid}", async (string uuid, AppDbContext db) =>
{
    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Content(RenderPage("<div class='card'><h3>设备不存在</h3><p>该 UUID 对应的设备不存在。</p></div>"), "text/html;charset=utf-8");

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
    return Results.Content(RenderPage(content, "home"), "text/html;charset=utf-8");
});

/// <summary>
/// GET /machine/{uuid}/task/{taskId} - 任务详情页：显示打卡排名和状态网格，支持 Web 打卡
/// </summary>
app.MapGet("/machine/{uuid}/task/{taskId}", async (string uuid, string taskId, AppDbContext db) =>
{
    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Content(RenderPage("<div class='card'><h3>设备不存在</h3></div>"), "text/html;charset=utf-8");

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
    return Results.Content(RenderPage(content, "home"), "text/html;charset=utf-8");
});

app.Run();
