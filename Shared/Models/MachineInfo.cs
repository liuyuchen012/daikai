namespace CheckIn.Shared.Models;

/// <summary>
/// 设备信息模型（用于服务器状态查询和 Web 面板展示）
/// </summary>
public class MachineInfo
{
    /// <summary>设备唯一标识符</summary>
    public string Uuid { get; set; } = string.Empty;
    /// <summary>设备名称</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>RSA 公钥（PEM 格式）</summary>
    public string PublicKey { get; set; } = string.Empty;
    /// <summary>最后在线时间</summary>
    public string? LastSeen { get; set; }
    /// <summary>当前是否在线</summary>
    public bool Online { get; set; }
}
