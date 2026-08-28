using System.IO;
using System.Text.Json;
using CheckIn.Client.Models;

namespace CheckIn.Client.Services;

/// <summary>
/// 课时划消本地存储服务，将数据持久化到 data/classhours.json
/// </summary>
public class ClassHourStore
{
    private static readonly JsonSerializerOptions _options = new() { PropertyNameCaseInsensitive = true, WriteIndented = true };
    private static readonly object _lock = new();
    private readonly string _filePath;

    public ClassHourStore()
    {
        var dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");
        Directory.CreateDirectory(dir);
        _filePath = Path.Combine(dir, "classhours.json");
    }

    /// <summary>加载数据，文件不存在时返回空数据；旧版本（v2 排课为字符串数组）自动迁移</summary>
    public ClassHourData Load()
    {
        lock (_lock)
        {
            try
            {
                if (!File.Exists(_filePath)) return new ClassHourData();
                var json = File.ReadAllText(_filePath);
                try
                {
                    var data = JsonSerializer.Deserialize<ClassHourData>(json, _options);
                    if (data != null) return data;
                }
                catch (JsonException)
                {
                    // 旧格式不兼容，走迁移
                }
                return MigrateLegacy(json);
            }
            catch
            {
                return new ClassHourData();
            }
        }
    }

    /// <summary>迁移旧版本数据（v2 排课 Dictionary&lt;string, List&lt;string&gt;&gt; → v3 带时间的排课条目）</summary>
    private static ClassHourData MigrateLegacy(string json)
    {
        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            var data = new ClassHourData { Version = 3 };

            if (root.TryGetProperty("Students", out var st) && st.ValueKind == JsonValueKind.Array)
                foreach (var el in st.EnumerateArray())
                {
                    var s = JsonSerializer.Deserialize<ChStudent>(el.GetRawText(), _options);
                    if (s != null) data.Students.Add(s);
                }

            if (root.TryGetProperty("Records", out var rc) && rc.ValueKind == JsonValueKind.Array)
                foreach (var el in rc.EnumerateArray())
                {
                    var r = JsonSerializer.Deserialize<ChRecord>(el.GetRawText(), _options);
                    if (r != null) data.Records.Add(r);
                }

            if (root.TryGetProperty("OffDays", out var od) && od.ValueKind == JsonValueKind.Array)
                foreach (var el in od.EnumerateArray())
                    if (el.ValueKind == JsonValueKind.String)
                        data.OffDays.Add(el.GetString()!);

            if (root.TryGetProperty("Schedule", out var sch) && sch.ValueKind == JsonValueKind.Object)
                foreach (var prop in sch.EnumerateObject())
                {
                    var list = new List<ScheduleEntry>();
                    if (prop.Value.ValueKind == JsonValueKind.Array)
                        foreach (var el in prop.Value.EnumerateArray())
                        {
                            if (el.ValueKind == JsonValueKind.String)
                                list.Add(new ScheduleEntry { StudentId = el.GetString()! });
                            else
                            {
                                var e = JsonSerializer.Deserialize<ScheduleEntry>(el.GetRawText(), _options);
                                if (e != null) list.Add(e);
                            }
                        }
                    data.Schedule[prop.Name] = list;
                }

            if (root.TryGetProperty("HoursPerHour", out var h) && h.ValueKind == JsonValueKind.Number)
                data.HoursPerHour = h.GetDouble();
            if (root.TryGetProperty("AutoDeduct", out var a) && a.ValueKind is JsonValueKind.True or JsonValueKind.False)
                data.AutoDeduct = a.GetBoolean();

            return data;
        }
        catch
        {
            return new ClassHourData();
        }
    }

    /// <summary>保存数据</summary>
    public void Save(ClassHourData data)
    {
        lock (_lock)
        {
            var json = JsonSerializer.Serialize(data, _options);
            File.WriteAllText(_filePath, json);
        }
    }
}
