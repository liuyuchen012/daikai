namespace CallServer.Models;

/// <summary>
/// 呼叫记录
/// 呼叫类型：urgent 紧急呼叫 / notice 普通通知 / speech 仅朗读（不弹窗）
/// </summary>
public class CallRecord
{
    public int Id { get; set; }

    public string Type { get; set; } = "notice";

    /// <summary>呼叫标题</summary>
    public string Title { get; set; } = "";

    /// <summary>呼叫内容</summary>
    public string Message { get; set; } = "";

    /// <summary>发送人（呼出端填写）</summary>
    public string Sender { get; set; } = "";

    /// <summary>目标设备 UUID；null 表示广播给全部设备</summary>
    public string? TargetUuid { get; set; }

    public DateTime CreatedAt { get; set; } = DateTime.Now;

    /// <summary>被控端确认时间（null 表示尚未送达）</summary>
    public DateTime? AckedAt { get; set; }

    /// <summary>确认送达的设备 UUID</summary>
    public string? AckedByUuid { get; set; }
}
