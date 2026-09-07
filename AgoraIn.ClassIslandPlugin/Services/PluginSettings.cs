using System.Text.Json;

namespace AgoraIn.ClassIslandPlugin.Services;

/// <summary>
/// 插件设置（持久化到插件配置目录 settings.json）
/// 可手动填写，或从本机 AgoraIn 大屏客户端的 config.json / client_uuid.txt 自动探测
/// </summary>
public class PluginSettings
{
    /// <summary>集控平台地址，如 http://192.168.1.100:5250</summary>
    public string ServerUrl { get; set; } = "";

    /// <summary>集控平台连接密码（与 AgoraIn 大屏客户端 config.json 的 ServerPassword 一致）</summary>
    public string Password { get; set; } = "";

    /// <summary>本设备 UUID（与 AgoraIn 大屏客户端 client_uuid.txt 一致）</summary>
    public string DeviceUuid { get; set; } = "";

    /// <summary>是否启用待下课时段通知（结合课表下课时间自动提醒）</summary>
    public bool PrenoticeEnabled { get; set; } = true;

    /// <summary>轮询间隔（秒）</summary>
    public int PollIntervalSeconds { get; set; } = 10;

    public static PluginSettings Load(string folder)
    {
        try
        {
            var path = Path.Combine(folder, "settings.json");
            if (File.Exists(path))
                return JsonSerializer.Deserialize<PluginSettings>(File.ReadAllText(path)) ?? new PluginSettings();
        }
        catch { }
        return new PluginSettings();
    }

    public void Save(string folder)
    {
        try
        {
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, "settings.json"),
                JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }

    /// <summary>
    /// 自动探测本机 AgoraIn 大屏客户端的配置：
    /// 读取 config.json（ServerIp / ServerPort / ServerPassword）与 client_uuid.txt（设备 UUID）
    /// 搜索：环境变量 AGORAIN_HOME 指定目录 → Program Files\AgoraIn → Program Files (x86)\AgoraIn → LocalAppData\Programs\AgoraIn
    /// </summary>
    public static void TryAutoDetect(PluginSettings s)
    {
        try
        {
            var candidates = new List<string>();
            var env = Environment.GetEnvironmentVariable("AGORAIN_HOME");
            if (!string.IsNullOrWhiteSpace(env))
                candidates.Add(env);
            candidates.AddRange(new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "AgoraIn"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "AgoraIn"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "AgoraIn"),
            });
            foreach (var dir in candidates)
            {
                var cfgPath = Path.Combine(dir, "config.json");
                var uuidPath = Path.Combine(dir, "client_uuid.txt");

                if (File.Exists(cfgPath))
                {
                    using var doc = JsonDocument.Parse(File.ReadAllText(cfgPath));
                    var root = doc.RootElement;
                    if (root.TryGetProperty("ServerIp", out var ip) && ip.ValueKind == JsonValueKind.String &&
                        root.TryGetProperty("ServerPort", out var port) && port.ValueKind == JsonValueKind.Number)
                    {
                        s.ServerUrl = $"http://{ip.GetString()}:{port.GetInt32()}";
                    }
                    if (root.TryGetProperty("ServerPassword", out var pwd) && pwd.ValueKind == JsonValueKind.String)
                        s.Password = pwd.GetString() ?? "";
                }
                if (File.Exists(uuidPath))
                    s.DeviceUuid = File.ReadAllText(uuidPath).Trim();

                if (!string.IsNullOrEmpty(s.ServerUrl) && !string.IsNullOrEmpty(s.Password) &&
                    !string.IsNullOrEmpty(s.DeviceUuid))
                    return;
                // 某项缺失时继续找下一个目录，但保留已找到的部分
            }
        }
        catch { }
    }
}
