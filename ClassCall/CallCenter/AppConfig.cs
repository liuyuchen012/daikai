using System.IO;
using System.Text.Json;

namespace CallCenter;

/// <summary>
/// 呼出端本地配置（服务端地址等），持久化到程序目录 config.json
/// </summary>
public class AppConfig
{
    /// <summary>服务端地址，如 http://192.168.1.100:5260</summary>
    public string ServerUrl { get; set; } = "http://127.0.0.1:5260";

    /// <summary>默认发送人</summary>
    public string Sender { get; set; } = "";

    public static string ConfigPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");

    public static AppConfig Load()
    {
        try
        {
            if (File.Exists(ConfigPath))
                return JsonSerializer.Deserialize<AppConfig>(File.ReadAllText(ConfigPath)) ?? new AppConfig();
        }
        catch { }
        return new AppConfig();
    }

    public void Save()
    {
        try
        {
            File.WriteAllText(ConfigPath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
    }
}
