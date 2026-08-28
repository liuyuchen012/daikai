using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;

namespace CheckIn.Client.Models;

/// <summary>课时划消数据模型（对应小程序 utils/classhours.js 的本地存储结构）</summary>
public class ClassHourData
{
    /// <summary>版本号，用于数据迁移（v3：排课细分时间 + 自动划消设置）</summary>
    public int Version { get; set; } = 3;

    /// <summary>学生列表</summary>
    public List<ChStudent> Students { get; set; } = new();

    /// <summary>课时记录（划消/增加历史）</summary>
    public List<ChRecord> Records { get; set; } = new();

    /// <summary>排课数据：日期(yyyy-MM-dd) -> 排课条目（含上课时间）</summary>
    public Dictionary<string, List<ScheduleEntry>> Schedule { get; set; } = new();

    /// <summary>不排课日集合（yyyy-MM-dd）</summary>
    public List<string> OffDays { get; set; } = new();

    /// <summary>设置：每小时上课消耗的课时数（支持小数，如 0.5 / 1 / 1.5）</summary>
    public double HoursPerHour { get; set; } = 1;

    /// <summary>设置：是否自动划消课时（课程结束时按时间表自动扣减）</summary>
    public bool AutoDeduct { get; set; } = false;
}

/// <summary>学生</summary>
public class ChStudent : INotifyPropertyChanged
{
    /// <summary>唯一ID</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString();

    /// <summary>姓名</summary>
    public string Name { get; set; } = "";

    private double _totalHours;
    /// <summary>总课时</summary>
    public double TotalHours
    {
        get => _totalHours;
        set { _totalHours = value; OnPropertyChanged(nameof(RemainingHours)); }
    }

    private double _usedHours;
    /// <summary>已划课时数</summary>
    public double UsedHours
    {
        get => _usedHours;
        set { _usedHours = value; OnPropertyChanged(nameof(RemainingHours)); }
    }

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

    /// <summary>当日排课的上课时间（UI 绑定用，不持久化）</summary>
    [JsonIgnore]
    public string ScheduleStartTime { get; set; } = "";

    /// <summary>当日排课的下课时间（UI 绑定用，不持久化）</summary>
    [JsonIgnore]
    public string ScheduleEndTime { get; set; } = "";

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

/// <summary>排课条目：学生 + 上课/下课时间（起止必填，无默认值，按实际时长计算课时）</summary>
public class ScheduleEntry
{
    /// <summary>学生ID</summary>
    public string StudentId { get; set; } = "";

    /// <summary>上课开始时间 HH:mm（必填，无默认值）</summary>
    public string StartTime { get; set; } = "";

    /// <summary>下课结束时间 HH:mm（必填，无默认值）</summary>
    public string EndTime { get; set; } = "";

    /// <summary>开始时间（解析失败返回占位值）</summary>
    [JsonIgnore]
    public DateTime Start => ParseTime(StartTime);

    /// <summary>结束时间；跨天课程（下课 &lt;= 上课）自动顺延一天</summary>
    [JsonIgnore]
    public DateTime End
    {
        get
        {
            var end = ParseTime(EndTime);
            return end <= Start ? end.AddDays(1) : end;
        }
    }

    /// <summary>起止时间是否均有效（空时间视为无效，不参与自动划消）</summary>
    [JsonIgnore]
    public bool IsValidTime => DateTime.TryParse(StartTime, out _) && DateTime.TryParse(EndTime, out _);

    /// <summary>课程时长（小时），起止时间无效时为 0</summary>
    [JsonIgnore]
    public double DurationHours => IsValidTime && End > Start ? (End - Start).TotalHours : 0;

    private static DateTime ParseTime(string time)
    {
        if (DateTime.TryParse(time, out var dt)) return dt;
        return DateTime.Today.AddHours(8);
    }
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

    /// <summary>自动划消去重键（yyyy-MM-dd|学生ID|上课时间），手动记录为空</summary>
    public string? SlotKey { get; set; }

    /// <summary>创建时间</summary>
    public string CreatedAt { get; set; } = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

    /// <summary>是否划消（负数为划消，UI 展示用）</summary>
    [JsonIgnore]
    public bool IsDeduct => Hours < 0;

    /// <summary>展示文本：+X / -X</summary>
    [JsonIgnore]
    public string HoursText => Hours < 0 ? $"{Hours:0.#}" : $"+{Hours:0.#}";
}
