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

    /// <summary>加载数据，文件不存在时返回空数据</summary>
    public ClassHourData Load()
    {
        lock (_lock)
        {
            try
            {
                if (!File.Exists(_filePath)) return new ClassHourData();
                var json = File.ReadAllText(_filePath);
                return JsonSerializer.Deserialize<ClassHourData>(json, _options) ?? new ClassHourData();
            }
            catch
            {
                return new ClassHourData();
            }
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
