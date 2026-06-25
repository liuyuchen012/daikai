using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using CheckIn.Server.Data;
using CheckIn.Shared.Models;
using Microsoft.EntityFrameworkCore;
using System.Web;

var builder = WebApplication.CreateBuilder(args);

var connStr = builder.Configuration.GetConnectionString("Default") ?? "Data Source=checkin.db";
builder.Services.AddDbContext<AppDbContext>(o => o.UseSqlite(connStr));
builder.Services.AddCors(o => o.AddDefaultPolicy(p => p.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader()));

var serverPassword = builder.Configuration.GetValue("ServerPassword", "admin123")!;
var serverName = builder.Configuration.GetValue("ServerName", "打卡中央控制平台")!;

var app = builder.Build();
app.UseCors();

using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();
}

// ---- Load HTML template ----
var templatePath = Path.Combine(builder.Environment.WebRootPath ?? "wwwroot", "template.html");
var templateContent = File.Exists(templatePath) ? File.ReadAllText(templatePath) : "<html><body>{CONTENT}</body></html>";

string RenderPage(string content)
{
    return templateContent
        .Replace("{TITLE}", HttpUtility.HtmlEncode(serverName))
        .Replace("{CONTENT}", content);
}

// ---- RSA helpers ----
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

string? GetPublicKey(AppDbContext db, string uuid) =>
    db.Machines.Where(m => m.Uuid == uuid).Select(m => m.PublicKey).FirstOrDefault();

static bool CheckPwd(JsonElement body, string expected)
{
    if (!body.TryGetProperty("password", out var p) || p.ValueKind != JsonValueKind.String)
        return false;
    return p.GetString() == expected;
}

// ===== API ENDPOINTS =====

app.MapGet("/api/status", async (AppDbContext db) =>
{
    var machines = await db.Machines.ToListAsync();
    var now = DateTime.Now;
    return machines.Select(m => new
    {
        uuid = m.Uuid,
        name = m.Name,
        online = m.LastSeen != null && (now - DateTime.Parse(m.LastSeen)).TotalSeconds < 300,
        last_seen = m.LastSeen
    });
});

app.MapPost("/api/register", async (AppDbContext db, JsonElement body) =>
{
    if (!CheckPwd(body, serverPassword))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var pubKey = body.GetProperty("public_key").GetString() ?? "";
    var name = body.GetProperty("name").GetString() ?? "";

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

app.MapPost("/api/sync_data", async (AppDbContext db, JsonElement body) =>
{
    if (!CheckPwd(body, serverPassword))
        return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("uuid").GetString() ?? "";
    var signature = body.GetProperty("signature").GetString() ?? "";
    var data = body.GetProperty("data").GetString() ?? "";
    var pubKey = GetPublicKey(db, uuid);
    if (pubKey == null) return Results.Json(new { error = "unknown machine" }, statusCode: 403);
    if (!VerifySignature(pubKey, data, signature))
        return Results.Json(new { error = "invalid signature" }, statusCode: 403);

    db.AttendanceRecords.Add(new AttendanceEntity { MachineUuid = uuid, Data = data, UpdatedAt = DateTime.Now.ToString("O") });
    var m = await db.Machines.FindAsync(uuid);
    if (m != null) m.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

app.MapPost("/api/load_data", async (AppDbContext db, JsonElement body) =>
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

    var latest = await db.AttendanceRecords.Where(a => a.MachineUuid == uuid).OrderByDescending(a => a.UpdatedAt).FirstOrDefaultAsync();
    var m = await db.Machines.FindAsync(uuid);
    if (m != null) m.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();

    var data = latest != null ? JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latest.Data) : new();
    return Results.Json(new { data });
});

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

app.MapPost("/api/update_machine_config", async (AppDbContext db, JsonElement body) =>
{
    var pwd = body.GetProperty("password").GetString() ?? "";
    if (pwd != serverPassword) return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("machine_uuid").GetString() ?? "";
    var machine = await db.Machines.FindAsync(uuid);
    if (machine != null) { machine.Config = body.GetProperty("config").GetRawText(); await db.SaveChangesAsync(); }
    return Results.Json(new { status = "ok" });
});

app.MapPost("/api/clear_attendance", async (AppDbContext db, JsonElement body) =>
{
    var pwd = body.GetProperty("password").GetString() ?? "";
    if (pwd != serverPassword) return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("machine_uuid").GetString() ?? "";
    db.AttendanceRecords.RemoveRange(db.AttendanceRecords.Where(a => a.MachineUuid == uuid));
    db.AttendanceRecords.Add(new AttendanceEntity { MachineUuid = uuid, Data = "{}", UpdatedAt = DateTime.Now.ToString("O") });
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

app.MapPost("/api/web_punch", async (AppDbContext db, JsonElement body) =>
{
    var pwd = body.GetProperty("password").GetString() ?? "";
    if (pwd != serverPassword) return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("machine_uuid").GetString() ?? "";
    var student = body.GetProperty("student_name").GetString() ?? "";

    var latest = await db.AttendanceRecords.Where(a => a.MachineUuid == uuid).OrderByDescending(a => a.UpdatedAt).FirstOrDefaultAsync();
    var data = latest != null ? JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latest.Data) ?? new() : new();

    if (!data.ContainsKey(student)) data[student] = new StudentAttendance { Name = student };
    if (data[student].FirstTime != null) return Results.Json(new { error = "该学生已经打卡" }, statusCode: 400);

    var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    data[student].FirstTime = now; data[student].Count++; data[student].History.Add(now);

    db.AttendanceRecords.Add(new AttendanceEntity { MachineUuid = uuid, Data = JsonSerializer.Serialize(data), UpdatedAt = DateTime.Now.ToString("O") });
    var m = await db.Machines.FindAsync(uuid);
    if (m != null) m.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

app.MapPost("/api/web_cancel_punch", async (AppDbContext db, JsonElement body) =>
{
    var pwd = body.GetProperty("password").GetString() ?? "";
    if (pwd != serverPassword) return Results.Json(new { error = "invalid password" }, statusCode: 403);

    var uuid = body.GetProperty("machine_uuid").GetString() ?? "";
    var student = body.GetProperty("student_name").GetString() ?? "";

    var latest = await db.AttendanceRecords.Where(a => a.MachineUuid == uuid).OrderByDescending(a => a.UpdatedAt).FirstOrDefaultAsync();
    if (latest == null) return Results.Json(new { error = "该机器无打卡数据" }, statusCode: 404);

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

    db.AttendanceRecords.Add(new AttendanceEntity { MachineUuid = uuid, Data = JsonSerializer.Serialize(data), UpdatedAt = DateTime.Now.ToString("O") });
    var m = await db.Machines.FindAsync(uuid);
    if (m != null) m.LastSeen = DateTime.Now.ToString("O");
    await db.SaveChangesAsync();
    return Results.Json(new { status = "ok" });
});

// ===== WEB PAGES =====

app.MapGet("/", async (AppDbContext db) =>
{
    var machines = await db.Machines.ToListAsync();
    var now = DateTime.Now;
    var rows = new StringBuilder();
    foreach (var m in machines)
    {
        var last = m.LastSeen != null ? DateTime.Parse(m.LastSeen) : (DateTime?)null;
        var online = last != null && (now - last.Value).TotalSeconds < 300;
        rows.Append($"<tr><td>{m.Uuid[..8]}...</td><td>{HttpUtility.HtmlEncode(m.Name)}</td><td class=\"{(online ? "status-online" : "status-offline")}\">{(online ? "在线" : "离线")}</td><td>{m.LastSeen}</td><td><a href=\"/machine/{m.Uuid}\" class=\"btn\">查看</a></td></tr>");
    }
    return Results.Content(RenderPage($"<h2>已注册机器</h2><table><tr><th>UUID</th><th>名称</th><th>状态</th><th>最后在线</th><th>操作</th></tr>{rows}</table>"), "text/html;charset=utf-8");
});

app.MapGet("/machine/{uuid}", async (string uuid, AppDbContext db) =>
{
    var machine = await db.Machines.FindAsync(uuid);
    if (machine == null) return Results.Content(RenderPage("<h2>机器不存在</h2>"), "text/html;charset=utf-8");

    var config = JsonSerializer.Deserialize<ClientConfig>(machine.Config) ?? new ClientConfig();
    var latest = await db.AttendanceRecords.Where(a => a.MachineUuid == uuid).OrderByDescending(a => a.UpdatedAt).FirstOrDefaultAsync();
    var data = latest != null ? JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(latest.Data) ?? new() : new();
    var updateTime = latest?.UpdatedAt ?? "从未同步";

    var punched = data.Values.Where(d => d.FirstTime != null).OrderBy(d => d.FirstTime).ToList();
    var rankRows = new StringBuilder();
    int i = 1;
    foreach (var p in punched) rankRows.Append($"<tr><td>{i++}</td><td>{HttpUtility.HtmlEncode(p.Name)}</td><td>{p.FirstTime}</td></tr>");
    var rankTable = rankRows.Length > 0 ? $"<table><tr><th>排名</th><th>姓名</th><th>打卡时间</th></tr>{rankRows}</table>" : "<p>暂无打卡记录</p>";

    var gridItems = new StringBuilder();
    foreach (var (name, sa) in data)
    {
        var cls = sa.FirstTime != null ? "punched" : "";
        var st = sa.FirstTime != null ? "true" : "false";
        gridItems.Append($"<div class=\"grid-item {cls}\" onclick=\"openPunchModal('{uuid}','{HttpUtility.HtmlEncode(name)}',{st})\">{HttpUtility.HtmlEncode(name)}</div>");
    }

    var configHtml = $"<li>学校：{HttpUtility.HtmlEncode(config.School ?? "未设置")}</li><li>年级：{HttpUtility.HtmlEncode(config.Nj ?? "未设置")}</li><li>班级：{HttpUtility.HtmlEncode(config.ClassId ?? "未设置")}</li><li>科目：{HttpUtility.HtmlEncode(config.Km ?? "未设置")}</li><li>网格行数：{config.Z}</li><li>网格列数：{config.L}</li>";

    var escapedConfig = HttpUtility.HtmlAttributeEncode(machine.Config);
    var content = $"<h2>机器详情 - {HttpUtility.HtmlEncode(machine.Name)}</h2><p>UUID: {uuid}</p><p>最后数据同步: {updateTime}</p><h3>当前配置</h3><ul>{configHtml}</ul><button class=\"btn\" onclick=\"openEditConfigModal('{uuid}','{escapedConfig}')\">编辑配置</button><button class=\"btn btn-danger\" onclick=\"openClearDataModal('{uuid}')\">清除打卡数据</button><h3>打卡排名 (最早打卡)</h3>{rankTable}<h3>学生打卡状态</h3><div class=\"grid\">{gridItems}</div><a href=\"/\" class=\"btn\">返回</a>";
    return Results.Content(RenderPage(content), "text/html;charset=utf-8");
});

app.Run();
