namespace CheckIn.Shared.Models;

/// <summary>
/// 学生打卡数据模型（Client 和 Server 之间传输的学生打卡状态）
/// </summary>
public class StudentAttendance
{
    /// <summary>学生姓名</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>打卡累计次数</summary>
    public int Count { get; set; }
    /// <summary>首次打卡时间（yyyy-MM-dd HH:mm:ss 格式），null 表示未打卡</summary>
    public string? FirstTime { get; set; }
    /// <summary>打卡历史记录列表（每次打卡的时间戳）</summary>
    public List<string> History { get; set; } = new();
}
