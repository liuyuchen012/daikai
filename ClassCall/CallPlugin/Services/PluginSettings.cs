using System.Text.Json;

namespace CallPlugin.Services;

/// <summary>
/// 插件设置（持久化到插件配置目录 settings.json）。
/// 首次使用时通过服务端注册接口自动换取 UUID / 密码。
/// </summary>
public class PluginSettings
{
    /// <summary>服务端地址，如 http://192.168.1.100:5260</summary>
    public string ServerUrl { get; set; } = "";

    /// <summary>设备名称（注册用，默认取计算机名；建议填写教室名，如「301 教室」）</summary>
    public string DeviceName { get; set; } = "";

    /// <summary>所属房间/位置</summary>
    public string Room { get; set; } = "";

    /// <summary>设备凭证（注册后由服务端下发）</summary>
    public string DeviceUuid { get; set; } = "";

    /// <summary>设备密码（注册后由服务端下发）</summary>
    public string Password { get; set; } = "";

    /// <summary>轮询间隔（秒）</summary>
    public int PollIntervalSeconds { get; set; } = 5;

    /// <summary>是否启用朗读（被控端收到呼叫后用 ClassIsland 朗读内容）</summary>
    public bool SpeechEnabled { get; set; } = true;

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

    public bool IsConfigured =>
        !string.IsNullOrEmpty(ServerUrl) && !string.IsNullOrEmpty(DeviceUuid) && !string.IsNullOrEmpty(Password);
}
