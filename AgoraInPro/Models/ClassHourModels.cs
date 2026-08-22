using System.Text.Json.Serialization;

namespace CheckIn.Client.Models;

/// <summary>课时划消数据模型（对应小程序 utils/classhours.js 的本地存储结构）</summary>
public class ClassHourData
{
    /// <summary>版本号，用于数据迁移</summary>
    public int Version { get; set; } = 2;

    /// <summary>学生列表</summary>
    public List<ChStudent> Students { get; set; } = new();

    /// <summary>课时记录（划消/增加历史）</summary>
    public List<ChRecord> Records { get; set; } = new();

    /// <summary>排课数据：日期(yyyy-MM-dd) -> 学生ID列表</summary>
    public Dictionary<string, List<string>> Schedule { get; set; } = new();

    /// <summary>不排课日集合（yyyy-MM-dd）</summary>
    public List<string> OffDays { get; set; } = new();
}

/// <summary>学生</summary>
public class ChStudent
{
    /// <summary>唯一ID</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>姓名</summary>
    public string Name { get; set; } = "";

    /// <summary>总课时</summary>
    public double TotalHours { get; set; } = 0;

    /// <summary>已划课时数</summary>
    public double UsedHours { get; set; } = 0;

    /// <summary>备注</summary>
    public string Remark { get; set; } = "";

    /// <summary>创建时间</summary>
    public string CreatedAt { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

    /// <summary>剩余课时（只读计算属性）</summary>
    [JsonIgnore]
    public double RemainingHours => TotalHours - UsedHours;

    /// <summary>是否已排入当前选中日期（UI 绑定用，不持久化）</summary>
    [JsonIgnore]
    public bool IsInSchedule { get; set; }
}

/// <summary>课时划消记录</summary>
public class ChRecord
{
    /// <summary>唯一ID</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>关联学生ID</summary>
    public string StudentId { get; set; } = "";

    /// <summary>日期</summary>
    public string Date { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");

    /// <summary>课时数（划消为正、增加为负）</summary>
    public double Hours { get; set; } = 0;

    /// <summary>备注/原因</summary>
    public string Note { get; set; } = "";

    /// <summary>创建时间</summary>
    public string CreatedAt { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

    /// <summary>是否划消（负数为划消，UI 展示用）</summary>
    [JsonIgnore]
    public bool IsDeduct => Hours < 0;

    /// <summary>展示文本：+X / -X</summary>
    [JsonIgnore]
    public string HoursText => Hours < 0 ? $"{Hours:0.#}" : $"+{Hours:0.#}";
}
