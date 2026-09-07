namespace CallCenter.Models;

/// <summary>被控设备信息（来自服务端 GET /api/devices）</summary>
public class DeviceInfo
{
    public int Id { get; set; }
    public string Uuid { get; set; } = "";
    public string Name { get; set; } = "";
    public string Room { get; set; } = "";
    public bool IsOnline { get; set; }
    public DateTime LastHeartbeat { get; set; }

    /// <summary>列表显示文本：名称（房间）</summary>
    public string DisplayName => string.IsNullOrWhiteSpace(Room) ? Name : $"{Name}（{Room}）";

    /// <summary>在线状态文本</summary>
    public string StatusText => IsOnline ? "在线" : "离线";

    /// <summary>最近心跳时间</summary>
    public string HeartbeatText => LastHeartbeat.ToString("MM-dd HH:mm");
}
