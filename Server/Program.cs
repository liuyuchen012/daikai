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

// ---- 密码哈希辅助方法（加盐 SHA256） ----
string Sha256(string input)
{
    var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(input));
    return Convert.ToHexString(bytes).ToLower();
}

/// <summary>生成加盐密码哈希，格式: salt:SHA256(salt+password)</summary>
string HashPassword(string password)
{
    var salt = Convert.ToHexString(RandomNumberGenerator.GetBytes(16)).ToLower();
    var hash = Sha256(salt + password);
    return $"{salt}:{hash}";
}

/// <summary>验证密码哈希，兼容旧版无盐 SHA256 格式</summary>
bool VerifyPassword(string password, string storedHash)
{
    if (string.IsNullOrEmpty(storedHash)) return false;
    // 新格式: salt:hash
    var idx = storedHash.IndexOf(':');
    if (idx > 0)
    {
        var salt = storedHash[..idx];
        var hash = storedHash[(idx + 1)..];
        return Sha256(salt + password) == hash;
    }
    // 旧格式: 纯 SHA256（向后兼容）
    return Sha256(password) == storedHash;
}

// ---- 加载服务器配置文件 config.json ----
var configPath = Path.Combine(AppContext.BaseDirectory, "config.json");
if (!File.Exists(configPath))
    configPath = Path.Combine(Directory.GetCurrentDirectory(), "config.json");

var configJson = File.Exists(configPath) ? JsonDocument.Parse(File.ReadAllText(configPath)).RootElement : default;
var cfgPort = configJson.TryGetProperty("Port", out var p) ? p.GetInt32() : 5250;
var serverName = configJson.TryGetProperty("ServerName", out var sn) ? sn.GetString() ?? "AgoraIn 集控平台" : "AgoraIn 集控平台";

// 服务器密码：必须通过 config.json 设置，不存在则生成随机密码并回写
string serverPassword;
if (configJson.TryGetProperty("ServerPassword", out var sp) && sp.ValueKind == JsonValueKind.String && !string.IsNullOrEmpty(sp.GetString()))
{
    serverPassword = sp.GetString()!;
}
else
{
    serverPassword = Convert.ToHexString(System.Security.Cryptography.RandomNumberGenerator.GetBytes(12)).ToLower();
    Console.WriteLine($"[安全] 未配置 ServerPassword，已自动生成: {serverPassword}");
    // 回写到 config.json
    try
    {
        var configObj = File.Exists(configPath)
            ? JsonSerializer.Deserialize<Dictionary<string, object>>(File.ReadAllText(configPath)) ?? new()
            : new();
        configObj["ServerPassword"] = serverPassword;
        File.WriteAllText(configPath, JsonSerializer.Serialize(configObj, new JsonSerializerOptions { WriteIndented = true }));
    }
    catch (Exception ex) { Console.WriteLine($"[警告] 无法回写 config.json: {ex.Message}"); }
}

// 会话密钥持久化（服务器重启不会导致所有用户重新登录）
var sessionSecretFile = Path.Combine(AppContext.BaseDirectory, "session_secret.txt");
string sessionSecret;
if (File.Exists(sessionSecretFile))
{
    sessionSecret = File.ReadAllText(sessionSecretFile).Trim();
}
else
{
    sessionSecret = Convert.ToHexString(System.Security.Cryptography.RandomNumberGenerator.GetBytes(32)).ToLower();
    File.WriteAllText(sessionSecretFile, sessionSecret);
}

// ---- 调试模式配置（生产环境请设为 false） ----
var debugMode = configJson.ValueKind != JsonValueKind.Undefined &&
                configJson.TryGetProperty("DebugMode", out var dm) &&
                dm.ValueKind == JsonValueKind.True && dm.GetBoolean();
if (debugMode)
    Console.WriteLine("[调试] DebugMode 已启用 - 可通过 X-Debug-Auth 头绕过鉴权");
Console.WriteLine();

var builder = WebApplication.CreateBuilder(args);
builder.WebHost.UseUrls($"http://0.0.0.0:{cfgPort}");

var connStr = builder.Configuration.GetConnectionString("Default") ?? "Data Source=checkin.db";
builder.Services.AddDbContext<AppDbContext>(o => o.UseSqlite(connStr));
builder.Services.AddCors(o => o.AddDefaultPolicy(p => p
    .SetIsOriginAllowed(origin => new Uri(origin).IsLoopback)  // 仅允许本地回环地址
    .AllowAnyMethod().AllowAnyHeader().AllowCredentials()));

// ---- 版本常量（供 /api/version、启动横幅与更新检查器使用） ----
const string ServerVersion = "3.2.4";
const string LatestClientVersion = "v3.2.4";
const string ClientDownloadUrl = "https://github.com/liuyuchen012/AgoraIn/releases";

// ---- 注册集控平台版本更新检查器（后台定时检查 GitHub 最新发布） ----
builder.Services.AddSingleton(sp => new ServerUpdateChecker(
    sp.GetRequiredService<ILogger<ServerUpdateChecker>>(), ServerVersion, ClientDownloadUrl));
builder.Services.AddHostedService(sp => sp.GetRequiredService<ServerUpdateChecker>());

var app = builder.Build();
app.UseCors();
var serverUpdateChecker = app.Services.GetRequiredService<ServerUpdateChecker>();

using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();

    // 确保新表/新列存在：EF 的 EnsureCreated 不会为“已存在的旧库”补建缺失的表与列，
    // 这里显式补建（幂等），否则旧库会因缺少 SignInTasks 等表导致迁移与查询失败。
    try
    {
        await EnsureSchemaAsync(db);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[WARN] 数据库结构补建失败（不影响服务启动）：{ex}");
    }

    // ===== 兼容旧版数据库（旧版仅有 Machines + AttendanceRecords，且无 TaskId 列、无 SignInTasks 表） =====
    // 迁移失败不应阻断服务器启动，记录日志后继续（不影响已有功能）
    try
    {
        await MigrateLegacyDataAsync(db);
    }
    catch (Exception ex)
    {
        Console.WriteLine($"[WARN] 旧数据兼容迁移失败（不影响服务启动）：{ex}");
    }

/// <summary>
/// 旧版数据兼容迁移（幂等，可重复执行）：
/// 1) 为 AttendanceRecords 补充 TaskId 列（旧数据置为 'default'），避免新代码查询 TaskId 报错；
/// 2) 将旧的打卡记录（按 设备+任务 分组）与机器配置中的 PendingTasks 回填为 SignInTaskEntity，
///    使“旧平台创建的签到任务”在新版中可被识别与展示；
/// 3) 合并同名设备：旧版“一台设备 = 一个任务”，导致一台设备的多个任务被存为多台设备，
///    此处按设备名称合并为一台，并迁移其任务与打卡数据。
/// </summary>
async Task MigrateLegacyDataAsync(AppDbContext db)
{
    // 注意：TaskId 列与 SignInTasks/Users/DeviceAssignments 表已由 EnsureSchemaAsync 补齐，这里不再处理。

    // 1) 回填：旧 AttendanceRecords（按 设备+任务 分组）转为 SignInTaskEntity
    foreach (var grp in await db.AttendanceRecords
                 .GroupBy(a => new { a.MachineUuid, a.TaskId }).ToListAsync())
    {
        var latest = grp.OrderByDescending(a => a.UpdatedAt).First();
        var data = JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latest.Data) ?? new();
        var shortCode = grp.Key.TaskId.StartsWith("signin_", StringComparison.OrdinalIgnoreCase)
            ? grp.Key.TaskId["signin_".Length..]
            : LegacyShortCode(db);
        if (await db.SignInTasks.AnyAsync(s => s.ShortCode == shortCode)) continue;

        var studentNames = data.Keys.ToList();
        var signInRecords = data.Where(kv => kv.Value.FirstTime != null)
            .Select(kv => new SignInRecord { Name = kv.Key, Time = kv.Value.FirstTime! }).ToList();
        var machine = await db.Machines.FindAsync(grp.Key.MachineUuid);
        db.SignInTasks.Add(new SignInTaskEntity
        {
            ShortCode = shortCode,
            MachineUuid = grp.Key.MachineUuid,
            Password = "",
            Classroom = "",
            Subject = machine?.Name ?? "默认任务",
            TaskName = machine?.Name ?? "默认任务",
            StudentList = JsonSerializer.Serialize(studentNames),
            SignInRecords = JsonSerializer.Serialize(signInRecords),
            CreatedAt = latest.UpdatedAt,
            Status = "active"
        });
    }
    await db.SaveChangesAsync();

    // 2b) 回填：machine.Config.PendingTasks 中尚未建表的任务
    foreach (var machine in await db.Machines.ToListAsync())
    {
        var cfg = JsonSerializer.Deserialize<ClientConfig>(machine.Config) ?? new();
        if (cfg.PendingTasks == null) continue;
        foreach (var pt in cfg.PendingTasks)
        {
            if (string.IsNullOrEmpty(pt.ShortCode)) continue;
            if (await db.SignInTasks.AnyAsync(s => s.ShortCode == pt.ShortCode)) continue;
            db.SignInTasks.Add(new SignInTaskEntity
            {
                ShortCode = pt.ShortCode,
                MachineUuid = machine.Uuid,
                Password = pt.Password,
                Classroom = pt.Classroom,
                Subject = pt.Subject,
                TaskName = pt.TaskName,
                StudentList = JsonSerializer.Serialize(pt.Students),
                SignInRecords = "[]",
                CreatedAt = DateTime.Now.ToString("O"),
                Status = "active"
            });
        }
    }
    await db.SaveChangesAsync();

    // 3) 合并同名设备：将同名设备的任务与打卡数据迁移到最近活跃的一台，其余删除
    // 注：GroupBy 后的 Where(g => g.Count() > 1) 无法翻译为 SQL，先取回内存再分组
    var machines = await db.Machines.ToListAsync();
    var dupGroups = machines.GroupBy(m => m.Name).Where(g => g.Count() > 1).ToList();
    foreach (var g in dupGroups)
    {
        var canonical = g.OrderByDescending(m => m.LastSeen).First();
        foreach (var dup in g.Where(m => m.Uuid != canonical.Uuid).ToList())
        {
            await db.AttendanceRecords.Where(a => a.MachineUuid == dup.Uuid)
                .ExecuteUpdateAsync(a => a.SetProperty(x => x.MachineUuid, canonical.Uuid));
            await db.SignInTasks.Where(s => s.MachineUuid == dup.Uuid)
                .ExecuteUpdateAsync(s => s.SetProperty(x => x.MachineUuid, canonical.Uuid));
            await db.DeviceAssignments.Where(d => d.MachineUuid == dup.Uuid)
                .ExecuteUpdateAsync(d => d.SetProperty(x => x.MachineUuid, canonical.Uuid));
            db.Machines.Remove(dup);
        }
    }
    await db.SaveChangesAsync();
}

/// <summary>
/// 显式补齐缺失的表与列（幂等）。EF 的 EnsureCreated 在“数据库文件已存在”时不会补建模型中新增的表/列，
/// 旧版数据库（仅有 Machines + AttendanceRecords）缺少 SignInTasks / Users / DeviceAssignments 表以及
/// AttendanceRecords.TaskId 列，必须在迁移前补齐，否则后续读写会报“no such table/column”。
/// 列类型使用 SQLite 宽松亲和类型即可，EF 在 SQLite 下不做严格的运行时结构校验。
/// </summary>
async Task EnsureSchemaAsync(AppDbContext db)
{
    await db.Database.ExecuteSqlRawAsync("""
        CREATE TABLE IF NOT EXISTS "SignInTasks" (
            "Id" INTEGER PRIMARY KEY AUTOINCREMENT,
            "ShortCode" TEXT, "MachineUuid" TEXT, "Password" TEXT,
            "Classroom" TEXT, "Subject" TEXT, "TaskName" TEXT,
            "StudentList" TEXT, "SignInRecords" TEXT, "CreatedAt" TEXT, "Status" TEXT)
        """);
    await db.Database.ExecuteSqlRawAsync("""
        CREATE TABLE IF NOT EXISTS "Users" (
            "Id" INTEGER PRIMARY KEY AUTOINCREMENT,
            "Username" TEXT, "PasswordHash" TEXT, "Role" TEXT,
            "DisplayName" TEXT, "CreatedAt" TEXT, "IsActive" INTEGER)
        """);
    await db.Database.ExecuteSqlRawAsync("""
        CREATE TABLE IF NOT EXISTS "DeviceAssignments" (
            "Id" INTEGER PRIMARY KEY AUTOINCREMENT,
            "UserId" INTEGER, "MachineUuid" TEXT, "TaskId" TEXT,
            "AssignedBy" TEXT, "CreatedAt" TEXT)
        """);
    await db.Database.ExecuteSqlRawAsync("""CREATE UNIQUE INDEX IF NOT EXISTS "IX_SignInTasks_ShortCode" ON "SignInTasks" ("ShortCode")""");
    await db.Database.ExecuteSqlRawAsync("""CREATE INDEX IF NOT EXISTS "IX_SignInTasks_MachineUuid" ON "SignInTasks" ("MachineUuid")""");
    await db.Database.ExecuteSqlRawAsync("""CREATE UNIQUE INDEX IF NOT EXISTS "IX_Users_Username" ON "Users" ("Username")""");
    await db.Database.ExecuteSqlRawAsync("""CREATE INDEX IF NOT EXISTS "IX_DeviceAssignments_UserId" ON "DeviceAssignments" ("UserId")""");
    await db.Database.ExecuteSqlRawAsync("""CREATE INDEX IF NOT EXISTS "IX_DeviceAssignments_UserId_MachineUuid" ON "DeviceAssignments" ("UserId", "MachineUuid")""");
    await db.Database.ExecuteSqlRawAsync("""
        CREATE TABLE IF NOT EXISTS "Calls" (
            "Id" INTEGER PRIMARY KEY AUTOINCREMENT,
            "Type" TEXT, "MachineUuid" TEXT, "Title" TEXT, "Message" TEXT,
            "MinutesBefore" INTEGER, "StudentNames" TEXT, "Sender" TEXT,
            "CreatedAt" TEXT, "Status" TEXT, "ExpiresAt" TEXT)
        """);
    await db.Database.ExecuteSqlRawAsync("""CREATE INDEX IF NOT EXISTS "IX_Calls_MachineUuid_Status" ON "Calls" ("MachineUuid", "Status")""");

        // Machines 补充 ClientVersion 列（旧库无此列；新库已由 EnsureCreated 建好，无需执行）
    if (!await ColumnExistsAsync(db, "Machines", "ClientVersion"))
    {
        await db.Database.ExecuteSqlRawAsync("""ALTER TABLE "Machines" ADD COLUMN "ClientVersion" TEXT NULL""");
    }

// AttendanceRecords 补充 TaskId 列（旧库无此列；新库已由 EnsureCreated 建好，无需执行）
    // 先通过 PRAGMA 检查列是否存在，避免对已存在的列重复 ALTER 产生 fail 日志
    if (!await ColumnExistsAsync(db, "AttendanceRecords", "TaskId"))
    {
        await db.Database.ExecuteSqlRawAsync("""ALTER TABLE "AttendanceRecords" ADD COLUMN "TaskId" TEXT NOT NULL DEFAULT 'default'""");
    }
    await db.Database.ExecuteSqlRawAsync("""UPDATE "AttendanceRecords" SET "TaskId"='default' WHERE "TaskId" IS NULL OR "TaskId"=''""");
}

/// <summary>
/// 检查 SQLite 表中是否存在指定列（通过 PRAGMA table_info，幂等、无副作用）
/// </summary>
async Task<bool> ColumnExistsAsync(AppDbContext db, string table, string column)
{
    var conn = db.Database.GetDbConnection();
    var needClose = conn.State != System.Data.ConnectionState.Open;
    if (needClose) await conn.OpenAsync();
    try
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"PRAGMA table_info(\"{table}\")";
        using var reader = await cmd.ExecuteReaderAsync();
        while (await reader.ReadAsync())
        {
            if (string.Equals(Convert.ToString(reader["name"]), column, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }
    finally
    {
        if (needClose) await conn.CloseAsync();
    }
}

/// <summary>生成一个不重复的 6 位短链码（用于旧数据回填）</summary>
string LegacyShortCode(AppDbContext db)
{
    string code;
    do { code = GenerateShortCode(); }
    while (db.SignInTasks.Any(s => s.ShortCode == code));
    return code;
}

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
    // 调试模式：sw_session=debug 视为已认证
    if (HasDebugCookie(ctx)) return true;
    if (!ctx.Request.Cookies.TryGetValue(CookieName, out var token)) return false;
    return ValidateToken(token);
}

/// <summary>
/// 从会话令牌中提取用户名
/// </summary>
string? GetUsername(HttpContext ctx)
{
    // 调试模式
    var (debugUser, _) = GetDebugCookieUser(ctx);
    if (debugUser != null) return debugUser;

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
    // 调试模式
    var (_, debugRole) = GetDebugCookieUser(ctx);
    if (debugRole != null) return debugRole;

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

// ---- 新角色体系辅助方法 ----
// 角色体系：admin（管理员）、teacher（普通教师）、student（学生）、parent（家长）
// 兼容旧角色：operator → teacher, viewer → student

/// <summary>标准化角色名（兼容旧角色名）</summary>
string NormalizeRole(string role) => role switch
{
    "operator" => "teacher",
    "viewer" => "student",
    _ => role
};

/// <summary>有效的角色列表</summary>
string[] ValidRoles = new[] { "admin", "teacher", "student", "parent" };

/// <summary>检查是否管理员或教师角色（有管理权限）</summary>
bool IsAdminOrTeacher(string? role) =>
    role == "admin" || role == "teacher" || role == "operator";

// ---- Bearer Token 认证（用于移动端 API） ----
// 复用现有的 HMAC Token 机制，但通过 Authorization: Bearer <token> 头传递

/// <summary>
/// 从请求的 Authorization 头中提取并验证 Bearer Token
/// </summary>
(string? Username, string? Role, string? Error) ParseBearerToken(HttpContext ctx)
{
    // 调试模式：通过 X-Debug-Auth 头绕过 Bearer Token 验证
    if (debugMode)
    {
        var (debugUser, debugRole) = GetDebugUser(ctx);
        if (debugRole != null)
            return (debugUser, debugRole, null);
    }

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
/// 生成移动端登录令牌（与 Web Cookie 令牌格式相同，但作为 JSON 返回）
/// </summary>
string MakeLoginToken(string username, string role)
{
    return MakeToken(username, role);
}

// ---- 调试鉴权辅助（Debug Auth Helpers） ----
// 当 config.json 中 DebugMode=true 时生效。通过 X-Debug-Auth 请求头可绕过正常鉴权。
// 用法：
//   Cookie 认证绕过: 请求页面时带上 Cookie: sw_session=debug 即可
//   Bearer 认证绕过: Header X-Debug-Auth: admin（角色名）
//                     Header X-Debug-Auth: teacher1:teacher（用户名:角色）

/// <summary>
/// 从 X-Debug-Auth 头提取调试用户信息。仅在 DebugMode=true 时有效。
/// 格式: "role"（仅角色）或 "username:role"
/// </summary>
(string? Username, string? Role) GetDebugUser(HttpContext ctx)
{
    if (!debugMode) return (null, null);
    var debugHeader = ctx.Request.Headers["X-Debug-Auth"].FirstOrDefault();
    if (string.IsNullOrEmpty(debugHeader)) return (null, null);
    var parts = debugHeader.Split(':');
    if (parts.Length >= 2)
        return (parts[0], parts[1]);
    return ("debug", debugHeader.Trim());
}

/// <summary>
/// 检查是否有有效的调试 Cookie（sw_session=debug）。仅在 DebugMode=true 时有效。
/// </summary>
bool HasDebugCookie(HttpContext ctx)
{
    if (!debugMode) return false;
    ctx.Request.Cookies.TryGetValue(CookieName, out var token);
    return token == "debug";
}

/// <summary>
/// 根据调试 Cookie/Header 返回模拟用户信息
/// </summary>
(string? Username, string? Role) GetDebugCookieUser(HttpContext ctx)
{
    if (!HasDebugCookie(ctx)) return (null, null);
    // 同时检查 X-Debug-Auth 头以确定角色
    var (username, role) = GetDebugUser(ctx);
    return (username ?? "debug_admin", role ?? "admin");
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

    // 仅对管理员：若后台检测到集控平台有新版本，注入弹窗（每个会话仅提示一次）
    var updateModal = string.Empty;
    if (isAdminUser && serverUpdateChecker.Info.HasUpdate)
    {
        var latest = serverUpdateChecker.Info.LatestVersion;
        var url = serverUpdateChecker.Info.DownloadUrl;
        updateModal =
            "<div id='serverUpdateModal' class='modal-overlay'>" +
              "<div class='modal-box'>" +
                "<span class='modal-close' onclick=\"closeModal('serverUpdateModal');sessionStorage.setItem('serverUpdateDismissed','" + latest + "')\">&times;</span>" +
                "<h3>🎉 集控平台发现新版本</h3>" +
                "<p style='font-size:14px;color:var(--text-secondary);line-height:1.6;'>检测到集控管理平台有新版本 <strong>" + latest + "</strong>（当前 " + ServerVersion + "）。<br>建议管理员前往下载更新，以获得最新功能与修复。</p>" +
                "<div class='form-actions'>" +
                  "<a class='btn' href='" + url + "' target='_blank' onclick=\"sessionStorage.setItem('serverUpdateDismissed','" + latest + "')\">前往下载</a>" +
                  "<button class='btn btn-ghost' onclick=\"closeModal('serverUpdateModal');sessionStorage.setItem('serverUpdateDismissed','" + latest + "')\">稍后提醒</button>" +
                "</div>" +
              "</div>" +
            "</div>" +
            "<script>if(sessionStorage.getItem('serverUpdateDismissed')!=='" + latest + "'){document.getElementById('serverUpdateModal').style.display='block';}</script>";
    }

    return templateContent
        .Replace("{TITLE}", HttpUtility.HtmlEncode(serverName))
        .Replace("{NAV_HOME}", activeNav == "home" ? "active" : "")
        .Replace("{NAV_USERS}", activeNav == "users" ? "active" : "")
        .Replace("{NAV_PROFILE}", activeNav == "profile" ? "active" : "")
        .Replace("{USERS_VISIBLE}", isAdminUser ? "block" : "none")
        .Replace("{CURRENT_USER}", HttpUtility.HtmlEncode(username))
        .Replace("{CONTENT}", content)
        .Replace("{UPDATE_MODAL}", updateModal);
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
/// 验证请求中的密码是否与服务器密码一致。调试模式下接受 X-Debug-Auth 头绕过。
/// </summary>
bool CheckPwd(JsonElement body, string expected, HttpContext? ctx = null)
{
    // 调试模式：存在 X-Debug-Auth 头时跳过密码验证
    if (ctx != null && debugMode)
    {
        var debugHeader = ctx.Request.Headers["X-Debug-Auth"].FirstOrDefault();
        if (!string.IsNullOrEmpty(debugHeader))
            return true;
    }

    if (!body.TryGetProperty("password", out var p) || p.ValueKind != JsonValueKind.String)
        return false;
    return p.GetString() == expected;
}

// ===== 调试 API 端点（仅在 DebugMode=true 时可用） =====

/// <summary>
/// GET /api/debug/status - 查看调试模式状态和可用角色
/// </summary>
app.MapGet("/api/debug/status", () =>
{
    if (!debugMode)
        return Results.Json(new { debug_mode = false, message = "调试模式未启用，请在 config.json 中设置 DebugMode: true" });

    return Results.Json(new
    {
        debug_mode = true,
        message = "调试模式已启用，以下方式可绕过鉴权：",
        usage = new
        {
            cookie = "设置 Cookie: sw_session=debug，同时可选 Header X-Debug-Auth: admin/teacher/student/parent",
            bearer = "设置 Header X-Debug-Auth: admin（仅角色）或 X-Debug-Auth: username:role",
            token = "GET /api/debug/token?role=admin 生成标准 Bearer Token",
            roles = new[] { "admin", "teacher", "student", "parent" }
        }
    });
});

/// <summary>
/// GET /api/debug/token - 生成一个有效的 Bearer Token（调试用）
/// Query: ?role=admin 指定角色（默认 admin）
/// Query: ?username=test&role=teacher 指定用户名和角色
/// </summary>
app.MapGet("/api/debug/token", (HttpContext ctx) =>
{
    if (!debugMode)
        return Results.Json(new { error = "调试模式未启用" }, statusCode: 403);

    var role = ctx.Request.Query["role"].FirstOrDefault() ?? "admin";
    var username = ctx.Request.Query["username"].FirstOrDefault() ?? $"debug_{role}";

    // 验证角色有效性
    var validRoles = new[] { "admin", "teacher", "student", "parent" };
    if (!validRoles.Contains(role))
        return Results.Json(new { error = $"无效角色: {role}，可用角色: {string.Join(", ", validRoles)}" }, statusCode: 400);

    var token = MakeToken(username, role);
    return Results.Json(new
    {
        token,
        user = new { username, role, display_name = $"调试{role}" },
        usage = "使用 Authorization: Bearer <token> 头，或 X-Debug-Auth: <role> 头"
    });
});

/// <summary>
/// POST /api/debug/login - 模拟任意用户登录，返回 Bearer Token
/// Body: { "role": "admin", "username": "optional" }
/// </summary>
app.MapPost("/api/debug/login", (JsonElement body) =>
{
    if (!debugMode)
        return Results.Json(new { error = "调试模式未启用" }, statusCode: 403);

    var role = body.TryGetProperty("role", out var r) ? r.GetString()?.Trim() ?? "admin" : "admin";
    var username = body.TryGetProperty("username", out var un) ? un.GetString()?.Trim() : null;
    username = string.IsNullOrEmpty(username) ? $"debug_{role}" : username;

    var validRoles = new[] { "admin", "teacher", "student", "parent" };
    if (!validRoles.Contains(role))
        return Results.Json(new { error = $"无效角色: {role}，可用角色: {string.Join(", ", validRoles)}" }, statusCode: 400);

    var token = MakeLoginToken(username, role);
    return Results.Json(new
    {
        token,
        user = new { username, role, display_name = $"调试{role}" }
    });
});

// ===== API 端点 =====

/// <summary>
/// GET /api/status - 获取所有已注册设备列表（含在线状态和任务数量）
/// </summary>
app.MapGet("/api/status", async (AppDbContext db) =>
{
    var machines = await db.Machines.ToListAsync();
    // BUG FIX: 任务数以 SignInTasks 为准，避免无打卡记录的任务被漏算
    var signInTasks = await db.SignInTasks.ToListAsync();
    var taskLookup = signInTasks
        .GroupBy(s => s.MachineUuid)
        .ToDictionary(g => g.Key, g => g.Select(x => $"signin_{x.ShortCode}").ToList());
    var now = DateTime.Now;

    var groupedMachines = machines.Select(m =>
    {
        var tasks = taskLookup.TryGetValue(m.Uuid, out var t) ? t : new List<string>();
        var last = m.LastSeen != null ? DateTime.Parse(m.LastSeen) : (DateTime?)null;
        var online = last != null && (now - last.Value).TotalSeconds < 300;

        return new
        {
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
/// GET /api/version - 返回服务端版本与最新客户端版本，供客户端"检查更新"功能使用
/// </summary>
app.MapGet("/api/version", (ServerUpdateChecker checker) =>
{
    var info = checker.Info;
    return Results.Json(new
    {
        server_version = ServerVersion,
        latest_client_version = info.HasUpdate ? info.LatestVersion : LatestClientVersion,
        download_url = ClientDownloadUrl,
        server_update_available = info.HasUpdate,
        server_latest_version = info.LatestVersion
    });
});

/// <summary>
/// GET /api/server_update - 返回集控平台自身是否有新版本（后台定时检查 GitHub 的结果），
/// 供 WPF 客户端轮询并弹窗通知管理员用户
/// </summary>
app.MapGet("/api/server_update", (ServerUpdateChecker checker) =>
{
    var info = checker.Info;
    return Results.Json(new
    {
        has_update = info.HasUpdate,
        latest_version = info.LatestVersion,
        current_version = info.CurrentVersion,
        download_url = info.DownloadUrl,
        last_checked = info.LastChecked
    });
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
app.MapPost("/api/register", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    if (!CheckPwd(body, serverPassword, ctx))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var pubKey = body.GetProperty("public_key").GetString() ?? "";
    var name = body.GetProperty("name").GetString() ?? "";
    var taskId = body.TryGetProperty("task_id", out var tid) ? tid.GetString() ?? "default" : "default";

    var existing = await db.Machines.FirstOrDefaultAsync(m => m.PublicKey == pubKey);
    if (existing != null)
    {
        existing.LastSeen = DateTime.Now.ToString("O");
        existing.ClientVersion = body.TryGetProperty("client_version", out var cv) ? cv.GetString() ?? "" : "";
        await db.SaveChangesAsync();
        return Results.Json(new { uuid = existing.Uuid, existing = true });
    }

    var machine = new MachineEntity
    {
        Uuid = Guid.NewGuid().ToString(),
        Name = name,
        PublicKey = pubKey,
        LastSeen = DateTime.Now.ToString("O"),
        ClientVersion = body.TryGetProperty("client_version", out var cv2) ? cv2.GetString() ?? "" : ""
    };
    db.Machines.Add(machine);
    await db.SaveChangesAsync();
    return Results.Json(new { uuid = machine.Uuid, existing = false });
});

/// <summary>
/// POST /api/sync_data - 客户端同步打卡数据到服务器（需 RSA 签名验证）
/// </summary>
app.MapPost("/api/sync_data", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    if (!CheckPwd(body, serverPassword, ctx))
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
    if (m != null)
    {
        m.LastSeen = DateTime.Now.ToString("O");
        m.ClientVersion = body.TryGetProperty("client_version", out var cv) ? cv.GetString() ?? "" : m.ClientVersion ?? "";
    }
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

/// <summary>
/// POST /api/load_data - 客户端从服务器加载最新打卡数据（需 challenge-签名验证）
/// </summary>
app.MapPost("/api/load_data", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    if (!CheckPwd(body, serverPassword, ctx))
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
app.MapPost("/api/get_config", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    if (!CheckPwd(body, serverPassword, ctx))
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
app.MapPost("/api/update_config", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    if (!CheckPwd(body, serverPassword, ctx))
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
/// POST /api/delete_machine - Web 面板删除设备及其所有打卡数据和签到任务
/// </summary>
app.MapPost("/api/delete_machine", async (AppDbContext db, JsonElement body) =>
{
    var pwd = body.GetProperty("password").GetString() ?? "";
    if (pwd != serverPassword) return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("machine_uuid").GetString() ?? "";
    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Json(new { error = "设备不存在" }, statusCode: 404);

    // 级联删除：打卡记录 + 签到任务 + 设备分配 + 设备本身
    db.AttendanceRecords.RemoveRange(db.AttendanceRecords.Where(a => a.MachineUuid == uuid));
    db.SignInTasks.RemoveRange(db.SignInTasks.Where(s => s.MachineUuid == uuid));
    db.DeviceAssignments.RemoveRange(db.DeviceAssignments.Where(d => d.MachineUuid == uuid));
    db.Machines.Remove(machine);
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

/// <summary>
/// POST /api/web_rename_machine - Web 面板修改设备名称（需已登录会话）
/// </summary>
app.MapPost("/api/web_rename_machine", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    if (!IsAuthenticated(ctx)) return Results.Json(new { error = "未登录" }, statusCode: 401);

    var uuid = body.GetProperty("machine_uuid").GetString() ?? "";
    var name = body.GetProperty("name").GetString() ?? "";
    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Json(new { error = "设备不存在" }, statusCode: 404);
    if (string.IsNullOrWhiteSpace(name)) return Results.Json(new { error = "设备名称不能为空" }, statusCode: 400);
    machine.Name = name.Trim();
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
    if (!ValidRoles.Contains(role))
        return Results.Json(new { error = "无效的角色，必须是 admin、teacher、student 或 parent" }, statusCode: 400);

    if (await db.Users.AnyAsync(u => u.Username == username))
        return Results.Json(new { error = "用户名已存在" }, statusCode: 409);

    var user = new UserEntity
    {
        Username = username,
        PasswordHash = HashPassword(password),
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
            var role = r.GetString() ?? "student";
            if (ValidRoles.Contains(role))
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
        if (!VerifyPassword(oldPassword, targetUser.PasswordHash))
            return Results.Json(new { error = "旧密码错误" }, statusCode: 403);
    }

    targetUser.PasswordHash = HashPassword(newPassword);
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
app.MapPost("/api/create_signin", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    if (!CheckPwd(body, serverPassword, ctx))
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

    // 向设备推送任务配置（PendingTasks），客户端轮询时自动拉取
    var deviceConfig = JsonSerializer.Deserialize<ClientConfig>(machine.Config) ?? new ClientConfig();
    deviceConfig.PendingTasks ??= new List<PendingTaskConfig>();
    deviceConfig.ConfigVersion++;
    deviceConfig.PendingTasks.Add(new PendingTaskConfig
    {
        ShortCode = shortCode,
        TaskId = taskId,
        Subject = subject,
        Classroom = classroom,
        TaskName = subject,
        Password = signPassword,
        Students = studentNames,
        CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
    });
    machine.Config = JsonSerializer.Serialize(deviceConfig);
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
app.MapPost("/api/signin_result", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    if (!CheckPwd(body, serverPassword, ctx))
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

    // 从数据库查询用户并验证密码（支持加盐哈希）
    var user = await db.Users.FirstOrDefaultAsync(u => u.Username == username && u.IsActive);
    if (user != null && !VerifyPassword(password, user.PasswordHash)) user = null;

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
    // 仅加载必要字段，避免全表 JSON 数据加载
    var taskGroups = await db.AttendanceRecords
        .Select(a => new { a.MachineUuid, a.TaskId })
        .Distinct()
        .ToListAsync();
    var taskLookup = taskGroups.GroupBy(x => x.MachineUuid)
        .ToDictionary(g => g.Key, g => g.Select(x => x.TaskId).Distinct().Count());
    var now = DateTime.Now;
    int onlineCount = 0, totalCount = machines.Count;

    var rows = new StringBuilder();
    foreach (var m in machines)
    {
        var last = m.LastSeen != null ? DateTime.Parse(m.LastSeen) : (DateTime?)null;
        var online = last != null && (now - last.Value).TotalSeconds < 300;
        if (online) onlineCount++;

        var taskCount = taskLookup.TryGetValue(m.Uuid, out var tc) ? tc : 0;
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
    var clientVersion = HttpUtility.HtmlEncode(string.IsNullOrEmpty(machine.ClientVersion) ? "未知" : machine.ClientVersion);
    // 重命名设备脚本（普通字符串，避免插值模板转义问题）
    var renameJs = "<script>function renameMachine() { var nn = prompt('请输入新的设备名称', ''); if (nn == null || nn.trim() == '') return; fetch('/api/web_rename_machine', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ machine_uuid: '" + uuid + "', name: nn.trim() }) }).then(function (r) { return r.json(); }).then(function (j) { if (j.status == 'ok') location.reload(); else alert(j.error || '重命名失败'); }).catch(function (e) { alert('重命名失败: ' + e); }); }</script>";

    var content = $@"<div class=""breadcrumb""><a href=""/"">设备总览</a> / {HttpUtility.HtmlEncode(machine.Name)}</div>
<div class=""page-header"">
<h2>{HttpUtility.HtmlEncode(machine.Name)}</h2>
<div style=""display:flex;gap:8px;"">
<span class=""badge {badgeCls}"">{badgeText}</span>
<span class=""badge"" style=""background:#eef2ff;color:#4338ca;"">客户端 {clientVersion}</span>
<button class=""btn btn-sm"" onclick=""renameMachine()"">重命名</button>
<button class=""btn btn-sm"" onclick=""openEditConfigModal('{uuid}','{escapedConfig}')"">编辑配置</button>
<button class=""btn btn-sm btn-danger"" onclick=""openDeleteMachineModal('{uuid}','{HttpUtility.HtmlEncode(machine.Name)}')"">删除设备</button>
</div>
</div>

<div class=""card"">
    <h3>任务列表</h3>
    <p style=""color:var(--text-secondary);font-size:13px;margin-bottom:16px;"">点击任务查看详细打卡数据</p>
    <div class=""task-grid"">{taskCards}</div>
</div>

<script>{renameJs}</script>

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
        var roleLabel = NormalizeRole(u.Role) switch
        {
            "admin" => "<span class='badge badge-online'>管理员</span>",
            "teacher" => "<span class='badge' style='background:#e8f0fe;color:#1967d2'>普通教师</span>",
            "student" => "<span class='badge badge-offline' style='background:#f3e8fd;color:#7c3aed'>学生</span>",
            "parent" => "<span class='badge' style='background:#fef7e0;color:#e37400'>家长</span>",
            _ => "<span class='badge badge-offline'>未知</span>"
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
        '<div class=""form-group""><label>角色</label><select name=""role""><option value=""student"">学生</option><option value=""parent"">家长</option><option value=""teacher"">普通教师</option><option value=""admin"">管理员</option></select></div>' +
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
        '<div class=""form-group""><label>角色</label><select name=""role""><option value=""student""' + (role==='student'?' selected':'') + '>学生</option><option value=""parent""' + (role==='parent'?' selected':'') + '>家长</option><option value=""teacher""' + (role==='teacher'?' selected':'') + '>普通教师</option><option value=""admin""' + (role==='admin'?' selected':'') + '>管理员</option></select></div>' +
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

    var roleLabel = NormalizeRole(user.Role) switch
    {
        "admin" => "管理员",
        "teacher" => "普通教师",
        "student" => "学生",
        "parent" => "家长",
        _ => "未知"
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

    var user = await db.Users.FirstOrDefaultAsync(u => u.Username == username);
    if (user == null || !VerifyPassword(password, user.PasswordHash))
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
/// POST /api/qrcode/generate - 管理员/教师创建签到任务，生成二维码数据（返回 shortCode）
/// 必须指定设备 UUID，任务与设备绑定。生成后向设备推送任务配置。
/// 需要 admin 或 teacher 角色
/// </summary>
app.MapPost("/api/qrcode/generate", async (AppDbContext db, HttpContext ctx) =>
{
    // Bearer Token 认证 + 权限检查
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    var normalizedRole = NormalizeRole(role ?? "");
    if (!IsAdminOrTeacher(normalizedRole))
        return Results.Json(new { error = "权限不足，仅管理员或教师可创建签到任务" }, statusCode: 403);

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
    if (string.IsNullOrEmpty(machineUuid))
        return Results.Json(new { error = "请选择设备" }, statusCode: 400);

    // 验证设备存在
    var targetMachine = await db.Machines.FindAsync(machineUuid);
    if (targetMachine == null)
        return Results.Json(new { error = "指定的设备不存在" }, statusCode: 404);

    // 解析学生名单
    List<string> studentNames;
    try { studentNames = JsonSerializer.Deserialize<List<string>>(studentNamesStr) ?? new(); }
    catch { return Results.Json(new { error = "学生名单格式错误" }, statusCode: 400); }

    // 生成唯一短链码
    string shortCode;
    do { shortCode = GenerateShortCode(); }
    while (await db.SignInTasks.AnyAsync(s => s.ShortCode == shortCode));

    var task = new SignInTaskEntity
    {
        ShortCode = shortCode,
        MachineUuid = machineUuid,
        Password = signPassword,
        Classroom = classroom,
        Subject = subject,
        TaskName = subject,
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
        MachineUuid = machineUuid,
        TaskId = taskId,
        Data = JsonSerializer.Serialize(initialData),
        UpdatedAt = DateTime.Now.ToString("O")
    });

    // 向设备推送任务配置：更新 MachineEntity.Config
    var deviceConfig = JsonSerializer.Deserialize<ClientConfig>(targetMachine.Config) ?? new ClientConfig();
    deviceConfig.PendingTasks ??= new List<PendingTaskConfig>();
    deviceConfig.ConfigVersion++;
    deviceConfig.PendingTasks.Add(new PendingTaskConfig
    {
        ShortCode = shortCode,
        TaskId = taskId,
        Subject = subject,
        Classroom = classroom,
        TaskName = subject,
        Password = signPassword,
        Students = studentNames,
        CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
    });
    targetMachine.Config = JsonSerializer.Serialize(deviceConfig);

    targetMachine.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();

    return Results.Json(new
    {
        short_code = shortCode,
        task_id = taskId,
        qrcode_url = $"/s/{shortCode}",
        subject,
        classroom,
        student_count = studentNames.Count,
        created_at = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
        machine_uuid = machineUuid,
        machine_name = targetMachine.Name
    });
});

/// <summary>
/// POST /api/qrcode/checkin - 学生/家长通过扫码签到（JSON API）
/// student 或 parent 角色均可使用
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
/// GET /api/mobile/dashboard - 管理员/教师仪表盘数据
/// 教师仅显示已分配设备的任务
/// </summary>
app.MapGet("/api/mobile/dashboard", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role))
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var machines = await db.Machines.ToListAsync();
    var signInTasks = await db.SignInTasks.ToListAsync();
    // 任务数以 SignInTasks 为准（与任务列表页一致），避免因无打卡记录而漏算
    var taskLookup = signInTasks
        .GroupBy(s => s.MachineUuid)
        .ToDictionary(g => g.Key, g => g.Count());
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
        task_count = taskLookup.TryGetValue(m.Uuid, out var tc) ? tc : 0,
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
/// 教师只能查看已分配设备的考勤数据
/// </summary>
app.MapGet("/api/mobile/attendance", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role))
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
/// GET /api/mobile/tasks - 获取签到任务列表（管理员/教师）
/// 教师只能看到已分配设备上的任务
/// </summary>
app.MapGet("/api/mobile/tasks", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role))
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
    if (!IsAdminOrTeacher(role))
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
    if (!IsAdminOrTeacher(role))
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
/// GET /api/mobile/students/history - 获取签到历史
/// 学生/家长：查看自己作为学生的签到记录
/// 管理员/教师：查看所有设备的签到任务汇总
/// </summary>
app.MapGet("/api/mobile/students/history", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);

    var normalizedRole = NormalizeRole(role ?? "");
    var history = new List<object>();

    if (normalizedRole == "admin" || normalizedRole == "teacher")
    {
        // 管理员/教师：汇总所有设备的签到任务
        var signInTasks = await db.SignInTasks.ToListAsync();
        foreach (var task in signInTasks)
        {
            var records = JsonSerializer.Deserialize<List<SignInRecord>>(task.SignInRecords) ?? new();
            var studentNames = JsonSerializer.Deserialize<List<string>>(task.StudentList) ?? new();
            history.Add(new
            {
                subject = task.Subject,
                classroom = task.Classroom,
                short_code = task.ShortCode,
                task_name = task.TaskName ?? task.Subject,
                student_count = studentNames.Count,
                signed_count = records.Count,
                status = task.Status,
                created_at = task.CreatedAt,
                records = records.Select(r => new
                {
                    name = r.Name,
                    time = r.Time
                }).ToList()
            });
        }

        return Results.Json(new
        {
            role = normalizedRole,
            total_tasks = history.Count,
            total_checkins = history.Sum(h => ((dynamic)h).signed_count),
            history = history.OrderByDescending(h => ((dynamic)h).created_at).ToList()
        });
    }
    else
    {
        // 学生/家长：查询该用户作为学生的签到记录
        var signInTasks = await db.SignInTasks.ToListAsync();
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
    }
});

// =============================================================================
// 移动端 - 设备管理 API
// =============================================================================

/// <summary>
/// GET /api/mobile/devices - 获取设备列表（管理员/教师）
/// </summary>
app.MapGet("/api/mobile/devices", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role))
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var machines = await db.Machines.ToListAsync();
    var now = DateTime.Now;

    // 教师获取其分配的设备列表
    if (role == "teacher")
    {
        var user = await db.Users.FirstOrDefaultAsync(u => u.Username == username);
        if (user == null) return Results.Json(new { devices = Array.Empty<object>() });

        var assignedUuids = await db.DeviceAssignments
            .Where(d => d.UserId == user.Id)
            .Select(d => d.MachineUuid)
            .Distinct()
            .ToListAsync();
        machines = machines.Where(m => assignedUuids.Contains(m.Uuid)).ToList();
    }

    var deviceList = machines.Select(m =>
    {
        var last = m.LastSeen != null ? DateTime.Parse(m.LastSeen) : (DateTime?)null;
        return new
        {
            uuid = m.Uuid,
            name = m.Name,
            online = last != null && (now - last.Value).TotalSeconds < 300,
            last_seen = m.LastSeen,
            client_version = m.ClientVersion ?? "",
            public_key = string.IsNullOrEmpty(m.PublicKey) ? "N/A" : m.PublicKey[..Math.Min(m.PublicKey.Length, 30)] + "..."
        };
    }).ToList();

    return Results.Json(new { devices = deviceList });
});

/// <summary>
/// POST /api/mobile/calls - 教师发送呼叫（JWT 鉴权）
/// 三种类型：prenotice（待下课时段通知）/ emergency（上课应急通知）/ summon（下课传唤）
/// </summary>
app.MapPost("/api/mobile/calls", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role))
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var machineUuid = body.GetProperty("machine_uuid").GetString() ?? "";
    if (string.IsNullOrEmpty(machineUuid))
        return Results.Json(new { error = "machine_uuid 不能为空" }, statusCode: 400);

    var machine = await db.Machines.FindAsync(machineUuid);
    if (machine == null)
        return Results.Json(new { error = "设备不存在" }, statusCode: 404);

    var type = body.GetProperty("type").GetString() ?? "prenotice";
    if (type is not ("prenotice" or "emergency" or "summon"))
        return Results.Json(new { error = "type 必须为 prenotice / emergency / summon" }, statusCode: 400);

    var title = body.GetProperty("title").GetString() ?? "";
    if (string.IsNullOrWhiteSpace(title))
        return Results.Json(new { error = "标题不能为空" }, statusCode: 400);

    var minutes = body.TryGetProperty("minutes_before", out var mb) ? mb.GetInt32() : 0;
    var call = new CallEntity
    {
        Type = type,
        MachineUuid = machineUuid,
        Title = title,
        Message = body.GetProperty("message").GetString() ?? "",
        MinutesBefore = Math.Max(0, minutes),
        StudentNames = body.TryGetProperty("student_names", out var sn) ? sn.GetString() ?? "" : "",
        Sender = username,
        Status = "pending",
        CreatedAt = DateTime.Now.ToString("O"),
        ExpiresAt = DateTime.Now.AddHours(2).ToString("O")
    };
    db.Calls.Add(call);
    await db.SaveChangesAsync();

    return Results.Json(new { id = call.Id, status = "ok" });
});

/// <summary>
/// GET /api/mobile/calls - 教师查询呼叫记录（JWT 鉴权），管理员可见全部、教师仅可见自己发送的
/// </summary>
app.MapGet("/api/mobile/calls", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role))
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var query = db.Calls.AsQueryable();
    if (role == "teacher")
        query = query.Where(c => c.Sender == username);

    var calls = await query.OrderByDescending(c => c.Id).Take(100).ToListAsync();
    return Results.Json(new
    {
        calls = calls.Select(c => new
        {
            id = c.Id,
            type = c.Type,
            machine_uuid = c.MachineUuid,
            title = c.Title,
            message = c.Message,
            minutes_before = c.MinutesBefore,
            student_names = c.StudentNames,
            sender = c.Sender,
            created_at = c.CreatedAt,
            status = c.Status
        })
    });
});

/// <summary>
/// DELETE /api/mobile/devices/{uuid} - 移动端删除设备（级联删除任务和打卡数据）
/// 仅管理员可操作
/// </summary>
app.MapDelete("/api/mobile/devices/{uuid}", async (string uuid, AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (role != "admin")
        return Results.Json(new { error = "权限不足，仅管理员可删除设备" }, statusCode: 403);

    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null)
        return Results.Json(new { error = "设备不存在" }, statusCode: 404);

    // 级联删除：打卡记录、签到任务、设备分配、设备
    db.AttendanceRecords.RemoveRange(db.AttendanceRecords.Where(a => a.MachineUuid == uuid));
    db.SignInTasks.RemoveRange(db.SignInTasks.Where(s => s.MachineUuid == uuid));
    db.DeviceAssignments.RemoveRange(db.DeviceAssignments.Where(d => d.MachineUuid == uuid));
    db.Machines.Remove(machine);
    await db.SaveChangesAsync();

    return Results.Json(new { status = "ok", message = "设备及其绑定的任务和签到数据已删除" });
});

/// <summary>
/// PUT /api/mobile/devices/{uuid}/rename - 修改设备名称
/// 管理员或已分配该设备的教师可操作。修改后推送到客户端配置。
/// </summary>
app.MapPut("/api/mobile/devices/{uuid}/rename", async (string uuid, AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role))
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    string bodyStr;
    using (var reader = new StreamReader(ctx.Request.Body, Encoding.UTF8))
        bodyStr = await reader.ReadToEndAsync();
    JsonElement body;
    try { body = JsonDocument.Parse(bodyStr).RootElement; }
    catch { return Results.Json(new { error = "无效的 JSON 格式" }, statusCode: 400); }

    var newName = body.TryGetProperty("name", out var n) ? n.GetString()?.Trim() ?? "" : "";
    if (string.IsNullOrEmpty(newName))
        return Results.Json(new { error = "设备名称不能为空" }, statusCode: 400);

    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null)
        return Results.Json(new { error = "设备不存在" }, statusCode: 404);

    machine.Name = newName;

    // 更新设备配置，推送名称变更到客户端
    var config = JsonSerializer.Deserialize<ClientConfig>(machine.Config) ?? new ClientConfig();
    config.DeviceName = newName;
    config.ConfigVersion++;
    machine.Config = JsonSerializer.Serialize(config);

    await db.SaveChangesAsync();

    return Results.Json(new { status = "ok", name = newName });
});

/// <summary>
/// PUT /api/mobile/tasks/{id}/rename - 修改任务名称
/// 管理员或教师可操作。修改后同时更新绑定的设备配置。
/// </summary>
app.MapPut("/api/mobile/tasks/{id}/rename", async (int id, AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role))
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    string bodyStr;
    using (var reader = new StreamReader(ctx.Request.Body, Encoding.UTF8))
        bodyStr = await reader.ReadToEndAsync();
    JsonElement body;
    try { body = JsonDocument.Parse(bodyStr).RootElement; }
    catch { return Results.Json(new { error = "无效的 JSON 格式" }, statusCode: 400); }

    var newName = body.TryGetProperty("name", out var n) ? n.GetString()?.Trim() ?? "" : "";
    if (string.IsNullOrEmpty(newName))
        return Results.Json(new { error = "任务名称不能为空" }, statusCode: 400);

    var task = await db.SignInTasks.FindAsync(id);
    if (task == null)
        return Results.Json(new { error = "任务不存在" }, statusCode: 404);

    task.TaskName = newName;
    task.Subject = newName; // 同步更新 subject

    // 推送名称变更到关联设备配置
    var machine = await db.Machines.FindAsync(task.MachineUuid);
    if (machine != null)
    {
        var config = JsonSerializer.Deserialize<ClientConfig>(machine.Config) ?? new ClientConfig();
        config.ConfigVersion++;

        // 更新 pending tasks 中的名称
        if (config.PendingTasks != null)
        {
            var pendingTask = config.PendingTasks.FirstOrDefault(t => t.ShortCode == task.ShortCode);
            if (pendingTask != null)
            {
                pendingTask.TaskName = newName;
                pendingTask.Subject = newName;
            }
        }
        machine.Config = JsonSerializer.Serialize(config);
    }

    await db.SaveChangesAsync();

    return Results.Json(new { status = "ok", name = newName });
});

// =============================================================================
// 移动端 - 普通任务管理 API（创建/删除设备上的普通签到任务）
// =============================================================================

/// <summary>
/// POST /api/mobile/devices/{uuid}/tasks - 为设备创建普通任务（推送任务配置到设备）
/// 管理员或已分配该设备的教师可操作
/// </summary>
app.MapPost("/api/mobile/devices/{uuid}/tasks", async (string uuid, AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null) return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role)) return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Json(new { error = "设备不存在" }, statusCode: 404);

    string bodyStr;
    using (var reader = new StreamReader(ctx.Request.Body, Encoding.UTF8)) bodyStr = await reader.ReadToEndAsync();
    JsonElement body;
    try { body = JsonDocument.Parse(bodyStr).RootElement; }
    catch { return Results.Json(new { error = "无效的 JSON 格式" }, statusCode: 400); }

    var subject = body.TryGetProperty("subject", out var s) ? s.GetString()?.Trim() ?? "" : "";
    var classroom = body.TryGetProperty("classroom", out var cr) ? cr.GetString()?.Trim() ?? "" : "";
    var taskName = body.TryGetProperty("task_name", out var tn) ? tn.GetString()?.Trim() ?? subject : subject;
    var password = body.TryGetProperty("password", out var pw) ? pw.GetString() ?? "" : "";
    var students = new List<string>();
    if (body.TryGetProperty("students", out var st) && st.ValueKind == JsonValueKind.Array)
        students = st.EnumerateArray().Select(e => e.GetString() ?? "").Where(n => !string.IsNullOrEmpty(n)).ToList();

    if (string.IsNullOrEmpty(subject) || string.IsNullOrEmpty(taskName))
        return Results.Json(new { error = "subject/task_name 不能为空" }, statusCode: 400);

    var cfg = JsonSerializer.Deserialize<ClientConfig>(machine.Config) ?? new ClientConfig();
    cfg.PendingTasks ??= new List<PendingTaskConfig>();
    var taskId = $"signin_{GenerateShortCode()}";
    cfg.PendingTasks.Add(new PendingTaskConfig
    {
        TaskId = taskId,
        ShortCode = taskId["signin_".Length..],
        Subject = subject,
        Classroom = classroom,
        TaskName = taskName,
        Password = password,
        Students = students,
        CreatedAt = DateTime.Now.ToString("O")
    });
    cfg.ConfigVersion++;
    machine.Config = JsonSerializer.Serialize(cfg);

    // 同步创建 SignInTaskEntity（使网页端可查看和管理该任务）
    db.SignInTasks.Add(new SignInTaskEntity
    {
        ShortCode = taskId["signin_".Length..],
        MachineUuid = uuid,
        Password = password,
        Classroom = classroom,
        Subject = subject,
        TaskName = taskName,
        StudentList = JsonSerializer.Serialize(students),
        SignInRecords = "[]",
        CreatedAt = DateTime.Now.ToString("O"),
        Status = "active"
    });

    // 同步创建 attendance 记录（使客户端能加载打卡数据）
    var initialData = new Dictionary<string, StudentAttendance>();
    foreach (var name in students)
        initialData[name] = new StudentAttendance { Name = name };
    db.AttendanceRecords.Add(new AttendanceEntity
    {
        MachineUuid = uuid,
        TaskId = taskId,
        Data = JsonSerializer.Serialize(initialData),
        UpdatedAt = DateTime.Now.ToString("O")
    });

    await db.SaveChangesAsync();

    return Results.Json(new { status = "ok", task_id = taskId, message = "任务已推送至设备" });
});

/// <summary>
/// DELETE /api/mobile/devices/{uuid}/tasks/{taskId} - 删除设备上的普通任务
/// 管理员或已分配该设备的教师可操作。删除设备的 PendingTasks 配置和对应打卡记录。
/// </summary>
app.MapDelete("/api/mobile/devices/{uuid}/tasks/{taskId}", async (string uuid, string taskId, AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null) return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role)) return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Json(new { error = "设备不存在" }, statusCode: 404);

    var cfg = JsonSerializer.Deserialize<ClientConfig>(machine.Config) ?? new ClientConfig();
    if (cfg.PendingTasks != null)
    {
        cfg.PendingTasks.RemoveAll(pt => pt.TaskId == taskId);
        cfg.ConfigVersion++;
        machine.Config = JsonSerializer.Serialize(cfg);
    }

    // 删除关联的打卡记录
    db.AttendanceRecords.RemoveRange(db.AttendanceRecords.Where(a => a.MachineUuid == uuid && a.TaskId == taskId));
    await db.SaveChangesAsync();

    return Results.Json(new { status = "ok", message = "任务已删除" });
});

/// <summary>
/// GET /api/mobile/tasks/{id}/qrcode - 查看任务的二维码链接
/// 管理员和教师均可查看
/// </summary>
app.MapGet("/api/mobile/tasks/{id}/qrcode", async (int id, AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role))
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var task = await db.SignInTasks.FindAsync(id);
    if (task == null)
        return Results.Json(new { error = "任务不存在" }, statusCode: 404);

    return Results.Json(new
    {
        id = task.Id,
        short_code = task.ShortCode,
        qrcode_url = $"/s/{task.ShortCode}",
        subject = task.Subject,
        classroom = task.Classroom,
        status = task.Status,
        created_at = task.CreatedAt,
        student_count = (JsonSerializer.Deserialize<List<string>>(task.StudentList) ?? new()).Count,
        signed_count = (JsonSerializer.Deserialize<List<SignInRecord>>(task.SignInRecords) ?? new()).Count
    });
});

// =============================================================================
// 设备分配 API（管理员将设备分配给教师，精细到任务级别）
// =============================================================================

/// <summary>
/// GET /api/mobile/assignments - 获取设备分配列表（管理员查看全部，教师查看自己的）
/// </summary>
app.MapGet("/api/mobile/assignments", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (!IsAdminOrTeacher(role))
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    List<DeviceAssignmentEntity> assignments;
    if (role == "admin")
    {
        assignments = await db.DeviceAssignments.OrderByDescending(d => d.CreatedAt).ToListAsync();
    }
    else
    {
        var user = await db.Users.FirstOrDefaultAsync(u => u.Username == username);
        if (user == null) return Results.Json(new { assignments = Array.Empty<object>() });
        assignments = await db.DeviceAssignments.Where(d => d.UserId == user.Id).ToListAsync();
    }

    var result = new List<object>();
    foreach (var a in assignments)
    {
        var teacher = await db.Users.FindAsync(a.UserId);
        var machine = await db.Machines.FindAsync(a.MachineUuid);
        result.Add(new
        {
            id = a.Id,
            user_id = a.UserId,
            teacher_name = teacher?.DisplayName ?? teacher?.Username ?? "未知",
            machine_uuid = a.MachineUuid,
            machine_name = machine?.Name ?? "未知",
            task_id = a.TaskId,
            task_name = a.TaskId != null
                ? (await db.SignInTasks.FirstOrDefaultAsync(s => $"signin_{s.ShortCode}" == a.TaskId))?.TaskName ?? a.TaskId
                : "所有任务",
            assigned_by = a.AssignedBy,
            created_at = a.CreatedAt
        });
    }

    return Results.Json(new { assignments = result });
});

/// <summary>
/// POST /api/mobile/assignments - 管理员分配设备给教师
/// task_id 为空表示分配该设备的所有任务
/// </summary>
app.MapPost("/api/mobile/assignments", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (role != "admin")
        return Results.Json(new { error = "权限不足，仅管理员可分配设备" }, statusCode: 403);

    string bodyStr;
    using (var reader = new StreamReader(ctx.Request.Body, Encoding.UTF8))
        bodyStr = await reader.ReadToEndAsync();
    JsonElement body;
    try { body = JsonDocument.Parse(bodyStr).RootElement; }
    catch { return Results.Json(new { error = "无效的 JSON 格式" }, statusCode: 400); }

    var teacherUserId = body.TryGetProperty("user_id", out var uid) && uid.TryGetInt32(out var tid) ? tid : 0;
    var machineUuid = body.TryGetProperty("machine_uuid", out var mu) ? mu.GetString()?.Trim() ?? "" : "";
    var taskId = body.TryGetProperty("task_id", out var tsk) ? tsk.GetString()?.Trim() : null;

    if (teacherUserId <= 0)
        return Results.Json(new { error = "请选择教师" }, statusCode: 400);
    if (string.IsNullOrEmpty(machineUuid))
        return Results.Json(new { error = "请选择设备" }, statusCode: 400);

    var teacher = await db.Users.FindAsync(teacherUserId);
    if (teacher == null || teacher.Role != "teacher")
        return Results.Json(new { error = "指定的用户不是教师角色" }, statusCode: 400);

    var machine = await db.Machines.FindAsync(machineUuid);
    if (machine == null)
        return Results.Json(new { error = "指定的设备不存在" }, statusCode: 404);

    // 检查是否已有相同分配
    var existing = await db.DeviceAssignments
        .FirstOrDefaultAsync(d => d.UserId == teacherUserId && d.MachineUuid == machineUuid && d.TaskId == taskId);
    if (existing != null)
        return Results.Json(new { error = "该分配已存在" }, statusCode: 409);

    var assignment = new DeviceAssignmentEntity
    {
        UserId = teacherUserId,
        MachineUuid = machineUuid,
        TaskId = taskId,
        AssignedBy = username ?? "admin",
        CreatedAt = DateTime.Now.ToString("O")
    };
    db.DeviceAssignments.Add(assignment);
    await db.SaveChangesAsync();

    return Results.Json(new
    {
        status = "ok",
        id = assignment.Id,
        message = taskId == null
            ? $"已将设备「{machine.Name}」的所有任务分配给教师「{teacher.DisplayName}」"
            : $"已将设备「{machine.Name}」的任务「{taskId}」分配给教师「{teacher.DisplayName}」"
    });
});

/// <summary>
/// DELETE /api/mobile/assignments/{id} - 删除设备分配
/// </summary>
app.MapDelete("/api/mobile/assignments/{id}", async (int id, AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (role != "admin")
        return Results.Json(new { error = "权限不足，仅管理员可管理分配" }, statusCode: 403);

    var assignment = await db.DeviceAssignments.FindAsync(id);
    if (assignment == null)
        return Results.Json(new { error = "分配记录不存在" }, statusCode: 404);

    db.DeviceAssignments.Remove(assignment);
    await db.SaveChangesAsync();

    return Results.Json(new { status = "ok", message = "分配已删除" });
});

/// <summary>
/// GET /api/mobile/teachers - 获取所有教师列表（用于分配设备时选择）
/// </summary>
app.MapGet("/api/mobile/teachers", async (AppDbContext db, HttpContext ctx) =>
{
    var (username, role, tokenError) = ParseBearerToken(ctx);
    if (tokenError != null)
        return Results.Json(new { error = tokenError }, statusCode: 401);
    if (role != "admin")
        return Results.Json(new { error = "权限不足" }, statusCode: 403);

    var teachers = await db.Users
        .Where(u => u.Role == "teacher" && u.IsActive)
        .Select(u => new
        {
            id = u.Id,
            username = u.Username,
            display_name = u.DisplayName
        })
        .ToListAsync();

    return Results.Json(new { teachers });
});

/// <summary>
/// POST /api/config_applied - 客户端确认已应用推送的配置任务
/// 客户端应用 PendingTasks 后调用此接口，服务端清除已推送任务
/// </summary>
app.MapPost("/api/config_applied", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    if (!CheckPwd(body, serverPassword, ctx))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("uuid").GetString() ?? "";
    var taskIds = body.TryGetProperty("applied_tasks", out var at)
        ? JsonSerializer.Deserialize<List<string>>(at.GetRawText()) ?? new()
        : new List<string>();

    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Json(new { error = "设备不存在" }, statusCode: 404);

    var config = JsonSerializer.Deserialize<ClientConfig>(machine.Config) ?? new ClientConfig();
    if (config.PendingTasks != null)
    {
        config.PendingTasks.RemoveAll(t => taskIds.Contains(t.TaskId));
        if (config.PendingTasks.Count == 0) config.PendingTasks = null;
        machine.Config = JsonSerializer.Serialize(config);
        await db.SaveChangesAsync();
    }

    return Results.Json(new { status = "ok" });
});

/// <summary>
/// POST /api/calls_pull - 设备拉取自己的待处理呼叫（CheckPwd 鉴权）
/// 返回该设备所有 pending 状态的呼叫，已过期的一并标记为 expired
/// </summary>
app.MapPost("/api/calls_pull", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    if (!CheckPwd(body, serverPassword, ctx))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("uuid").GetString() ?? "";
    if (string.IsNullOrEmpty(uuid)) return Results.Json(new { error = "uuid required" }, statusCode: 400);

    var now = DateTime.Now;
    // 先清理已过期的 pending 呼叫
    var expired = await db.Calls
        .Where(c => c.MachineUuid == uuid && c.Status == "pending")
        .ToListAsync();
    foreach (var e in expired)
    {
        if (DateTime.TryParse(e.ExpiresAt, out var exp) && exp < now)
            e.Status = "expired";
    }
    if (expired.Any(e => e.Status == "expired"))
        await db.SaveChangesAsync();

    var calls = await db.Calls
        .Where(c => c.MachineUuid == uuid && c.Status == "pending")
        .OrderBy(c => c.Id)
        .ToListAsync();

    return Results.Json(new
    {
        calls = calls.Select(c => new
        {
            id = c.Id,
            type = c.Type,
            title = c.Title,
            message = c.Message,
            minutes_before = c.MinutesBefore,
            student_names = c.StudentNames,
            sender = c.Sender,
            created_at = c.CreatedAt
        })
    });
});

/// <summary>
/// POST /api/calls_ack - 设备确认已收到并显示呼叫（CheckPwd 鉴权）
/// </summary>
app.MapPost("/api/calls_ack", async (AppDbContext db, JsonElement body, HttpContext ctx) =>
{
    if (!CheckPwd(body, serverPassword, ctx))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var id = body.GetProperty("id").GetInt32();
    var call = await db.Calls.FindAsync(id);
    if (call == null) return Results.Json(new { error = "not found" }, statusCode: 404);

    if (call.Status == "pending")
    {
        call.Status = "acknowledged";
        await db.SaveChangesAsync();
    }
    return Results.Json(new { status = "ok" });
});

/// <summary>
/// POST /api/client_update - 客户端查询最新客户端版本（GitHub Release 资产）
/// </summary>
static int CompareVersions(string a, string b)
{
    int[] P(string v) { var p = v.Trim().TrimStart('v', 'V').Split('.'); var n = new int[4]; for (var i = 0; i < 4; i++) if (i < p.Length && int.TryParse(p[i], out var x)) n[i] = x; return n; }
    var pa = P(a); var pb = P(b);
    for (var i = 0; i < 4; i++) { var c = pa[i].CompareTo(pb[i]); if (c != 0) return c; }
    return 0;
}

app.MapPost("/api/client_update", async (JsonElement body) =>
{
    if (!CheckPwd(body, serverPassword, null))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    try
    {
        using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(12) };
        http.DefaultRequestHeaders.Add("User-Agent", "AgoraIn-Client");
        var resp = await http.GetAsync("https://api.github.com/repos/liuyuchen012/AgoraIn/releases/latest");
        if (!resp.IsSuccessStatusCode) return Results.Json(new { has_update = false });
        using var doc = JsonDocument.Parse(await resp.Content.ReadAsStringAsync());
        var tag = doc.RootElement.GetProperty("tag_name").GetString() ?? "";
        if (string.IsNullOrEmpty(tag)) return Results.Json(new { has_update = false });

        // 优先取安装包，其次便携 zip
        string? downloadUrl = null;
        foreach (var a in doc.RootElement.GetProperty("assets").EnumerateArray())
        {
            var name = a.GetProperty("name").GetString() ?? "";
            if (name == "AgoraIn-Setup-v" + tag.TrimStart('v', 'V') + ".exe") { downloadUrl = a.GetProperty("browser_download_url").GetString(); break; }
        }
        if (downloadUrl == null)
        {
            foreach (var a in doc.RootElement.GetProperty("assets").EnumerateArray())
                if ((a.GetProperty("name").GetString() ?? "") == "Client.win-x64.zip") { downloadUrl = a.GetProperty("browser_download_url").GetString(); break; }
        }

        var latest = tag.StartsWith("v", StringComparison.OrdinalIgnoreCase) ? tag : "v" + tag;
        var hasUpdate = CompareVersions(latest, LatestClientVersion) > 0;
        return Results.Json(new { has_update = hasUpdate, latest_version = latest, download_url = downloadUrl ?? "" });
    }
    catch
    {
        return Results.Json(new { has_update = false });
    }
});

// ---- 启动横幅（命令行艺术字） ----
void PrintBanner()
{
    string ESC(int f, int b, int style) => $"\u001b[{style};{f};{b}m";
    var R = "\u001b[0m"; // 重置
    var art = new[]
    {
        ESC(37,40,0) + "    " + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀▀▀▀▀▀▀" + ESC(97,40,1) + "▄" + ESC(37,40,0) + "              " + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀" + ESC(97,40,1) + "▀▀▀▀▀" + ESC(97,47,1) + "▄▄" + ESC(90,47,1) + "▀" + ESC(37,40,0) + "▄" + ESC(90,40,1) + "▄" + ESC(37,40,0) + "     " + ESC(97,40,1) + "▄▄" + ESC(97,47,1) + "▀▀▀▀▀▄▄" + ESC(90,47,1) + "▀" + ESC(37,40,0) + "▄" + ESC(90,40,1) + "▄" + ESC(37,40,0) + "    " + ESC(97,40,1) + "█" + ESC(97,47,1) + "▀▀▀▀▀▀▀▀▀▀▄▄" + ESC(90,47,1) + "▀" + ESC(37,40,0) + "▄" + ESC(90,40,1) + "▄" + ESC(37,40,0) + "        " + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀▀▀▀▀▀▀" + ESC(97,40,1) + "▄" + ESC(37,40,0) + "          " + ESC(97,40,1) + "▄▄▄▄▄▄▄▄" + ESC(37,40,0) + "▄▄" + ESC(90,40,1) + "▄" + ESC(97,40,1) + "█" + ESC(97,47,1) + "▀▀▀▀▀▀▀▀▀▀▄▄" + ESC(90,47,1) + "▀" + ESC(37,40,0) + "▄" + ESC(90,40,1) + "▄" + ESC(37,40,0) + "    " + R,
        ESC(37,40,0) + "   " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + "▀  " + ESC(93,40,1) + "▄▄■" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(90,40,1) + "▄" + ESC(37,40,0) + "          " + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "▀        " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▀" + ESC(90,40,1) + "▄" + ESC(37,40,0) + "  " + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "▀   " + ESC(93,40,1) + "■" + ESC(37,40,0) + " " + ESC(93,40,1) + "▄▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "▀" + ESC(97,47,1) + "▄" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▀" + ESC(90,40,1) + "▄" + ESC(37,40,0) + "  " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█      " + ESC(93,40,1) + "■" + ESC(37,40,0) + " " + ESC(93,40,1) + "▄▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "▀" + ESC(97,47,1) + "▄" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▀" + ESC(90,40,1) + "▄" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + "▀  " + ESC(93,40,1) + "▄▄■" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(90,40,1) + "▄" + ESC(37,40,0) + "        " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█▀▀▀▀▀" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█      " + ESC(93,40,1) + "■" + ESC(37,40,0) + " " + ESC(93,40,1) + "▄▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "▀" + ESC(97,47,1) + "▄" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▀" + ESC(90,40,1) + "▄" + ESC(37,40,0) + "  " + R,
        ESC(37,40,0) + "   " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + " " + ESC(93,40,1) + "▄▀" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "        " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█  " + ESC(93,40,1) + "▄■·" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + "█" + ESC(90,40,1) + "█" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█         " + ESC(93,40,1) + "▀▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▀" + ESC(90,40,1) + "▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "▀" + ESC(97,47,1) + "▄" + ESC(37,40,0) + "▄         " + ESC(93,40,1) + "▀▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▀" + ESC(90,40,1) + "▄" + ESC(37,40,0) + "    " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + " " + ESC(93,40,1) + "▄▀" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "       " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█ " + ESC(93,40,1) + "▄■" + ESC(37,40,0) + "  " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█ " + ESC(31,40,0) + "░" + ESC(37,40,0) + "        " + ESC(93,40,1) + "▀▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▀" + ESC(90,40,1) + "▄" + ESC(37,40,0) + "  " + R,
        ESC(37,40,0) + "  " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌" + ESC(93,40,1) + "▐" + ESC(37,40,0) + "  " + ESC(97,40,1) + "▄" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "      " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌" + ESC(93,40,1) + "▄▀" + ESC(37,40,0) + "  " + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀▀" + ESC(97,40,1) + "▄▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▄" + ESC(90,40,1) + "▀" + ESC(37,40,0) + " " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌    " + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀" + ESC(97,40,1) + "▄" + ESC(37,40,0) + "▄   " + ESC(93,40,1) + "▌" + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▐" + ESC(90,40,1) + "▌" + ESC(37,40,0) + " " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌   " + ESC(97,40,1) + "█" + ESC(97,47,1) + "▀" + ESC(97,40,1) + "▄" + ESC(37,40,0) + "▄   " + ESC(93,40,1) + "▌" + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▐" + ESC(90,40,1) + "▌" + ESC(37,40,0) + "  " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌" + ESC(93,40,1) + "▐" + ESC(37,40,0) + "  " + ESC(97,40,1) + "▄" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "      " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(93,40,1) + "▐" + ESC(37,40,0) + "    " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(31,40,0) + "▒▒░░" + ESC(37,40,0) + " " + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀" + ESC(97,40,1) + "▄" + ESC(37,40,0) + "▄ " + ESC(31,40,0) + "░" + ESC(37,40,0) + " " + ESC(93,40,1) + "▌" + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▐" + ESC(90,40,1) + "▌" + R,
        ESC(37,40,0) + "  " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌" + ESC(93,40,1) + "│" + ESC(37,40,0) + " " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(97,40,1) + "█" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█ " + ESC(93,40,1) + "▌" + ESC(37,40,0) + "  " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,47,1) + "▄" + ESC(90,40,1) + "▀▀" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█    " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(37,40,0) + "  " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█   " + ESC(97,40,1) + "█" + ESC(97,47,1) + "▄" + ESC(97,40,1) + "▄▀" + ESC(37,40,0) + "   " + ESC(97,40,1) + "▄█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(37,40,0) + "   " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌" + ESC(93,40,1) + "│" + ESC(37,40,0) + " " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(97,40,1) + "█" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█  " + ESC(91,40,1) + "▄▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(31,40,0) + "▓▓▓▓" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(31,40,0) + "▒" + ESC(91,41,1) + "░" + ESC(31,40,0) + "▓▒░" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + R,
        ESC(37,40,0) + "  " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + " " + ESC(91,40,1) + "▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(97,40,1) + "▐▌" + ESC(91,40,1) + "▄" + ESC(37,40,0) + "  " + ESC(91,40,1) + "▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "    " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█    " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(97,40,1) + "█▀▀▀▀█" + ESC(37,40,0) + "▄  " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█    " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(37,40,0) + "  " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█ " + ESC(31,40,0) + "▓▒░" + ESC(37,40,0) + " " + ESC(31,40,0) + "░░" + ESC(37,40,0) + " " + ESC(97,40,1) + "▄█" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▄" + ESC(37,40,0) + "     " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + " " + ESC(91,40,1) + "▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(97,40,1) + "▐▌" + ESC(91,40,1) + "▄" + ESC(37,40,0) + "  " + ESC(91,40,1) + "▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "    " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█ " + ESC(91,40,1) + "▓" + ESC(91,41,1) + "▓▒" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(31,40,0) + "█" + ESC(91,41,1) + "░░" + ESC(31,40,0) + "█" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(91,41,1) + "░▒░" + ESC(31,40,0) + "█▓" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + R,
        ESC(37,40,0) + "  " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(91,40,1) + "▐" + ESC(91,41,1) + "▒" + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌▄" + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "▄" + ESC(91,40,1) + "▀▄" + ESC(91,41,1) + "▒" + ESC(91,40,1) + "▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "   " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█ " + ESC(91,40,1) + "▄██" + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(97,40,1) + "█▄" + ESC(37,40,0) + " " + ESC(93,40,1) + "■" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(90,47,1) + "░▒" + ESC(90,40,1) + "▄" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█  " + ESC(31,40,0) + "▄" + ESC(91,40,1) + "▄" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(97,47,1) + " " + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(31,40,0) + "▐" + ESC(91,40,1) + "██▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + "█" + ESC(90,40,1) + "█" + ESC(37,40,0) + " " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌" + ESC(31,40,0) + "█" + ESC(91,40,1) + "▄▌" + ESC(97,40,1) + "█" + ESC(97,43,1) + "▀" + ESC(97,40,1) + "▄" + ESC(31,40,0) + "▒" + ESC(91,41,1) + "░" + ESC(91,40,1) + "█" + ESC(97,43,1) + "▀" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▀" + ESC(90,40,1) + "▄" + ESC(37,40,0) + "    " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(91,40,1) + "▐" + ESC(91,41,1) + "▒" + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌▄" + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "▄" + ESC(91,40,1) + "▀▄" + ESC(91,41,1) + "▒" + ESC(91,40,1) + "▄" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "   " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█ " + ESC(91,41,1) + "▒░" + ESC(31,40,0) + "█" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(91,41,1) + "░░▒▓" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(97,47,1) + " " + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(91,41,1) + "░" + ESC(91,40,1) + "██" + ESC(31,40,0) + "██" + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + "█" + ESC(90,40,1) + "█" + R,
        ESC(37,40,0) + " " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌" + ESC(91,41,1) + "▒░" + ESC(31,40,0) + "▄" + ESC(97,40,1) + "▀" + ESC(37,40,0) + "█▀▀" + ESC(31,40,0) + "▄" + ESC(91,41,1) + "░▒░▒░" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "  " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█ " + ESC(91,40,1) + "░" + ESC(91,41,1) + "▒▓" + ESC(31,40,0) + "▌" + ESC(97,40,1) + "▀" + ESC(97,47,1) + "▄" + ESC(97,40,1) + "█" + ESC(31,40,0) + "▐▓▌" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▓" + ESC(90,40,1) + "█" + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌" + ESC(91,41,1) + "░▒▓▒" + ESC(97,40,1) + "▀" + ESC(97,47,1) + "▄▀" + ESC(37,40,0) + "▀" + ESC(31,40,0) + "█" + ESC(91,40,1) + "█▓" + ESC(91,41,1) + "▒" + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▐" + ESC(90,40,1) + "▌" + ESC(37,40,0) + " " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌" + ESC(91,41,1) + "░▓▒" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(91,41,1) + "▒" + ESC(91,40,1) + "█▓" + ESC(91,41,1) + "▒" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▐" + ESC(90,40,1) + "▌" + ESC(37,40,0) + "  " + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌" + ESC(91,41,1) + "▒░" + ESC(31,40,0) + "▄" + ESC(97,40,1) + "▀" + ESC(37,40,0) + "█▀▀" + ESC(31,40,0) + "▄" + ESC(91,41,1) + "░▒░▒░" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,47,1) + "█" + ESC(37,40,0) + "  " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█ " + ESC(31,40,0) + "██▒" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(91,41,1) + "░░▒▓" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(91,41,1) + "▒" + ESC(91,40,1) + "█▓" + ESC(91,41,1) + "▒" + ESC(31,40,0) + "█" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + R,
        ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "█ " + ESC(31,40,0) + "▒░▒█" + ESC(37,40,0) + " " + ESC(31,40,0) + "▀▀" + ESC(97,40,1) + "▄▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "▄" + ESC(31,40,0) + "▀█▒" + ESC(37,40,0) + " " + ESC(97,40,1) + "▀" + ESC(97,47,1) + "▄" + ESC(90,47,1) + "▀" + ESC(90,40,1) + "▄" + ESC(97,40,1) + "▐" + ESC(97,47,1) + "▌" + ESC(37,40,0) + "▌ " + ESC(91,40,1) + "░▒" + ESC(91,41,1) + "▒░" + ESC(31,40,0) + "▄▄▓▒▌" + ESC(97,40,1) + "█" + ESC(90,47,1) + "░▓" + ESC(90,40,1) + "█" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(31,40,0) + "█" + ESC(91,41,1) + "▓▒░" + ESC(31,40,0) + "▄▄" + ESC(37,40,0) + " " + ESC(31,40,0) + "▓" + ESC(91,41,1) + "░▒░" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(37,40,0) + "  " + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(91,41,1) + "░▓▒░" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(91,41,1) + "░▒░░" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(37,40,0) + " " + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "█ " + ESC(31,40,0) + "▒░▒█" + ESC(37,40,0) + " " + ESC(31,40,0) + "▀▀" + ESC(97,40,1) + "▄▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "▄" + ESC(31,40,0) + "▀█▒" + ESC(37,40,0) + " " + ESC(97,40,1) + "▀" + ESC(97,47,1) + "▄" + ESC(90,47,1) + "▀" + ESC(90,40,1) + "▄" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█ " + ESC(31,40,0) + "░░░" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(31,40,0) + "█" + ESC(37,40,0) + "█" + ESC(91,41,1) + "░▒▓▒" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(91,41,1) + "▓█▓▒░" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + R,
        ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + " " + ESC(31,40,0) + "░░" + ESC(37,40,0) + " " + ESC(97,40,1) + "▄▄" + ESC(97,47,1) + "▀▀▀" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▄" + ESC(97,47,1) + "▀" + ESC(97,40,1) + "▄" + ESC(37,40,0) + " " + ESC(31,40,0) + "▒░" + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀ " + ESC(90,47,1) + "▄" + ESC(90,40,1) + "▀" + ESC(37,40,0) + " " + ESC(97,40,1) + "▀" + ESC(97,47,1) + "▄" + ESC(37,40,0) + "▄ " + ESC(31,40,0) + "▀" + ESC(91,41,1) + "░" + ESC(31,40,0) + "██▓▒░" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(90,47,1) + "▓" + ESC(90,40,1) + "█" + ESC(37,40,0) + "  " + ESC(97,40,1) + "▀" + ESC(97,47,1) + "▄" + ESC(37,40,0) + "▄" + ESC(31,40,0) + "█▓▒░" + ESC(37,40,0) + " " + ESC(31,40,0) + "▒░▓▀" + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀ " + ESC(90,47,1) + "▄" + ESC(90,40,1) + "▀" + ESC(37,40,0) + " " + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "▀" + ESC(31,40,0) + "▄" + ESC(91,41,1) + "░" + ESC(31,40,0) + "█▓" + ESC(97,40,1) + "█" + ESC(37,40,0) + "█" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(31,40,0) + "░▓▀" + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀ " + ESC(90,47,1) + "▄" + ESC(90,40,1) + "▀" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + " " + ESC(31,40,0) + "░░" + ESC(37,40,0) + " " + ESC(97,40,1) + "▄▄" + ESC(97,47,1) + "▀▀▀" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▄" + ESC(97,47,1) + "▀" + ESC(97,40,1) + "▄" + ESC(37,40,0) + " " + ESC(31,40,0) + "▒░" + ESC(97,40,1) + "▄" + ESC(97,47,1) + "▀ " + ESC(90,47,1) + "▄" + ESC(90,40,1) + "▀" + ESC(97,40,1) + "█" + ESC(97,47,1) + " " + ESC(37,40,0) + " " + ESC(31,40,0) + "░░░" + ESC(37,40,0) + " " + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(31,40,0) + "█" + ESC(37,40,0) + "█" + ESC(91,41,1) + "░▒▓▓" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(91,41,1) + "░▒░" + ESC(31,40,0) + "█░" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + R,
        ESC(97,40,1) + "█" + ESC(97,47,1) + "▄" + ESC(97,40,1) + "▄▄▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▄▄" + ESC(90,40,1) + "▀▀▀" + ESC(37,40,0) + "   " + ESC(97,40,1) + "▀▄" + ESC(97,47,1) + "▀ " + ESC(90,47,1) + "▄" + ESC(90,40,1) + "▀" + ESC(37,40,0) + "     " + ESC(97,40,1) + "▀▀" + ESC(97,47,1) + "▄" + ESC(97,40,1) + "▄▄▄" + ESC(97,47,1) + "▀" + ESC(97,40,1) + "▄▄▄█" + ESC(90,47,1) + "░" + ESC(37,40,0) + "▀     " + ESC(97,40,1) + "▀▀" + ESC(97,47,1) + "▄" + ESC(97,40,1) + "▄▄▄▄▄" + ESC(97,47,1) + "▀▀" + ESC(90,47,1) + "▄" + ESC(37,40,0) + "▀" + ESC(90,40,1) + "▀" + ESC(37,40,0) + "   " + ESC(97,40,1) + "█" + ESC(97,47,1) + "▄" + ESC(97,40,1) + "▄▄▄▄▄▄█" + ESC(37,40,0) + "█" + ESC(97,40,1) + "█▄" + ESC(97,47,1) + "▀▀" + ESC(90,47,1) + "▄" + ESC(37,40,0) + "▀" + ESC(90,40,1) + "▀" + ESC(37,40,0) + "   " + ESC(97,40,1) + "█" + ESC(97,47,1) + "▄" + ESC(97,40,1) + "▄▄▄" + ESC(97,47,1) + "▀" + ESC(37,40,0) + "█" + ESC(90,47,1) + "▄▄" + ESC(90,40,1) + "▀▀▀" + ESC(37,40,0) + "   " + ESC(97,40,1) + "▀▄" + ESC(97,47,1) + "▀ " + ESC(90,47,1) + "▄" + ESC(90,40,1) + "▀" + ESC(37,40,0) + "  " + ESC(97,40,1) + "█" + ESC(97,47,1) + "▄" + ESC(97,40,1) + "▄▄▄▄▄█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(31,40,0) + "█" + ESC(97,47,1) + "▄" + ESC(97,41,1) + "▄▄▄▄" + ESC(97,40,1) + "█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + ESC(97,40,1) + "█" + ESC(97,41,1) + "▄▄▄▄" + ESC(97,40,1) + "▄█" + ESC(37,40,0) + "██" + ESC(90,40,1) + "█" + R,
    };
    foreach (var line in art) Console.WriteLine(line);
    Console.WriteLine($"\u001b[0;97;1;42m version={ServerVersion} \u001b[0m");
    Console.WriteLine();
}

PrintBanner();
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

// ---- 集控平台版本更新检查器：后台定时查询 GitHub 最新发布 ----
/// <summary>集控平台版本更新信息（由后台服务填充）</summary>
public class ServerUpdateInfo
{
    public bool HasUpdate { get; set; }
    public string LatestVersion { get; set; } = "";
    public string CurrentVersion { get; set; } = "";
    public string DownloadUrl { get; set; } = "";
    public DateTime LastChecked { get; set; }
}

/// <summary>
/// 后台托管服务：启动时及之后每 30 分钟检查一次 GitHub Releases 最新版本，
/// 与当前集控平台版本比较，发现新版本时写入 ServerUpdateInfo 供 Web 与客户端读取
/// </summary>
public class ServerUpdateChecker : IHostedService
{
    private readonly ILogger<ServerUpdateChecker> _logger;
    private readonly string _currentVersion;
    private readonly string _downloadUrl;
    public ServerUpdateInfo Info { get; }
    private Timer? _timer;

    public ServerUpdateChecker(ILogger<ServerUpdateChecker> logger, string currentVersion, string downloadUrl)
    {
        _logger = logger;
        _currentVersion = currentVersion;
        _downloadUrl = downloadUrl;
        Info = new ServerUpdateInfo { CurrentVersion = currentVersion, DownloadUrl = downloadUrl };
    }

    public Task StartAsync(CancellationToken cancellationToken)
    {
        _ = CheckNowAsync();
        _timer = new Timer(_ => _ = CheckNowAsync(), null, TimeSpan.FromMinutes(30), TimeSpan.FromMinutes(30));
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken)
    {
        _timer?.Dispose();
        return Task.CompletedTask;
    }

    /// <summary>立即向 GitHub 查询最新发布版本并更新状态</summary>
    public async Task CheckNowAsync()
    {
        try
        {
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
            http.DefaultRequestHeaders.Add("User-Agent", "AgoraIn-Server");
            var resp = await http.GetAsync("https://api.github.com/repos/liuyuchen012/AgoraIn/releases/latest");
            if (!resp.IsSuccessStatusCode) return;
            var json = await resp.Content.ReadFromJsonAsync<JsonElement>();
            var tag = json.GetProperty("tag_name").GetString() ?? "";
            if (string.IsNullOrEmpty(tag)) return;
            var latest = tag.StartsWith("v", StringComparison.OrdinalIgnoreCase) ? tag : "v" + tag;
            Info.LatestVersion = latest;
            Info.HasUpdate = CompareVersion(latest, _currentVersion) > 0;
            Info.LastChecked = DateTime.Now;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "检查 GitHub 版本更新失败");
        }
    }

    /// <summary>比较两个版本号，a &gt; b 返回 1，相等 0，否则 -1</summary>
    private static int CompareVersion(string a, string b)
    {
        var pa = ParseVersion(a);
        var pb = ParseVersion(b);
        for (var i = 0; i < 3; i++)
        {
            var c = pa[i].CompareTo(pb[i]);
            if (c != 0) return c;
        }
        return 0;
    }

    /// <summary>将版本字符串解析为 [major, minor, patch] 三个整数</summary>
    private static int[] ParseVersion(string v)
    {
        var parts = v.Trim().TrimStart('v', 'V').Split('.');
        var nums = new int[3];
        for (var i = 0; i < 3; i++)
            if (i < parts.Length && int.TryParse(parts[i], out var n)) nums[i] = n;
        return nums;
    }
}
