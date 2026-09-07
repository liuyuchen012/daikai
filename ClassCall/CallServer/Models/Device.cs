namespace CallServer.Models;

/// <summary>
/// 被控设备（教室一体机上运行的 ClassIsland 插件端）
/// </summary>
public class Device
{
    public int Id { get; set; }

    /// <summary>设备唯一标识（注册时生成）</summary>
    public string Uuid { get; set; } = Guid.NewGuid().ToString("N");

    /// <summary>设备凭证（拉取/确认呼叫时需要）</summary>
    public string Password { get; set; } = "";

    /// <summary>设备名称（如：301 教室）</summary>
    public string Name { get; set; } = "";

    /// <summary>所属位置/房间</summary>
    public string Room { get; set; } = "";

    /// <summary>最近一次心跳时间（超过 30 秒视为离线）</summary>
    public DateTime LastHeartbeat { get; set; } = DateTime.Now;

    public DateTime CreatedAt { get; set; } = DateTime.Now;

    /// <summary>是否在线（由心跳时间计算）</summary>
    public bool IsOnline => (DateTime.Now - LastHeartbeat).TotalSeconds <= 30;
}
