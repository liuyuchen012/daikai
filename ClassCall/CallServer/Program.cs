using CallServer.Data;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddDbContext<AppDbContext>(opt =>
    opt.UseSqlite(builder.Configuration.GetConnectionString("Default")
                  ?? "Data Source=classcall.db"));

// 允许呼出端/插件端跨域访问（局域网部署时前端可能使用不同端口）
builder.Services.AddCors(opt => opt.AddDefaultPolicy(p => p.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader()));

var app = builder.Build();

// 启动时自动建库
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<AppDbContext>();
    db.Database.EnsureCreated();
}

app.UseCors();
app.MapControllers();

// 默认监听本机所有网卡，端口可被 appsettings.json / 命令行覆盖
var url = builder.Configuration["Urls"] ?? "http://0.0.0.0:5260";
app.Urls.Clear();
app.Urls.Add(url);

Console.WriteLine($"[ClassCall] 服务端启动：{url}");

app.Run();
