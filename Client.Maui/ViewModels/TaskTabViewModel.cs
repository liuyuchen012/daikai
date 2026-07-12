using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Models;
using CheckIn.Client.Maui.Services;
using CheckIn.Shared.Models;

namespace CheckIn.Client.Maui.ViewModels;

/// <summary>
/// 每个标签页的独立配置，持久化到各任务目录下的 config.json
/// </summary>
public class TabConfig
{
    public string Name { get; set; } = "";
    public string Km { get; set; } = "";
    public int ButtonRows { get; set; } = 6;
    public int ButtonCols { get; set; } = 6;
    public bool OnlineMode { get; set; } = true;
    public bool IsSignInTask { get; set; } = false;
    public string? SignInTaskId { get; set; }
}

/// <summary>
/// 标签页信息（持久化到 workspace.json）
/// </summary>
public class TabInfo
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
}

/// <summary>
/// 每个打卡任务标签页的 ViewModel
/// </summary>
public class TaskTabViewModel : INotifyPropertyChanged, IDisposable
{
    public string TabId { get; }
    public TabConfig Config { get; } = new();

    private readonly string _tabDir;
    private readonly string _dataFile;
    private readonly string _nameFile;
    private readonly string _configFile;
    private ServerService? _server;
    private CancellationTokenSource? _syncCts;

    private string _globalServerIp = "";
    private int _globalServerPort = 5250;
    private string _globalServerPassword = "";

    /// <summary>学生数据集合</summary>
    public ObservableCollection<StudentModel> Students { get; } = new();
    /// <summary>打卡排名集合</summary>
    public ObservableCollection<RankingItem> Ranking { get; } = new();

    private bool _isOnline;
    public bool IsOnline
    {
        get => _isOnline;
        set { _isOnline = value; OnPropertyChanged(); OnPropertyChanged(nameof(StatusText)); OnPropertyChanged(nameof(StatusColor)); }
    }

    private string _statusMessage = "就绪";
    public string StatusMessage { get => _statusMessage; set { _statusMessage = value; OnPropertyChanged(); } }

    private string _lastCheck = "从未检查";
    public string LastCheck { get => _lastCheck; set { _lastCheck = value; OnPropertyChanged(); } }

    public string StatusText => IsOnline ? "服务器: 在线" : "服务器: 离线";
    public string StatusColor => IsOnline ? "#34a853" : "#ea4335";

    private int _totalStudents;
    public int TotalStudents { get => _totalStudents; set { _totalStudents = value; OnPropertyChanged(); OnPropertyChanged(nameof(PunchInfo)); } }

    private int _punchedCount;
    public int PunchedCount { get => _punchedCount; set { _punchedCount = value; OnPropertyChanged(); OnPropertyChanged(nameof(PunchInfo)); OnPropertyChanged(nameof(PunchPercent)); } }

    public string PunchInfo => $"总人数: {TotalStudents}  |  已打卡: {PunchedCount}  ({PunchPercent})";
    public string PunchPercent => TotalStudents > 0 ? $"{PunchedCount * 100.0 / TotalStudents:F1}%" : "0%";

    public string TabDisplayName => string.IsNullOrEmpty(Config.Km) ? Config.Name : $"{Config.Name} - {Config.Km}";

    // ---- Commands ----
    public ICommand CheckInCommand { get; } = null!;
    public ICommand CancelCheckInCommand { get; } = null!;
    public ICommand ExportCommand { get; } = null!;
    public ICommand ImportCommand { get; } = null!;
    public ICommand ClearAllCommand { get; } = null!;
    public ICommand ShowAdminSettingsCommand { get; } = null!;
    public ICommand CheckServerStatusCommand { get; } = null!;
    public ICommand LoadFromServerCommand { get; } = null!;
    public ICommand SyncToServerCommand { get; } = null!;

    private readonly ServerService? _sharedServer;

    public static TaskTabViewModel CreateBackgroundInstance(string tabId, string baseDir, AppConfig globalConfig, ServerService serverService)
    {
        var tab = new TaskTabViewModel(tabId, baseDir, globalConfig, serverService, true);
        return tab;
    }

    public TaskTabViewModel(string tabId, string baseDir, AppConfig globalConfig, ServerService? serverService = null, bool backgroundMode = false)
    {
        TabId = tabId;
        _tabDir = Path.Combine(baseDir, "data", "tabs", tabId);
        Directory.CreateDirectory(_tabDir);
        _dataFile = Path.Combine(_tabDir, "attendance.dat");
        _nameFile = Path.Combine(_tabDir, "name.txt");
        _configFile = Path.Combine(_tabDir, "config.json");

        _sharedServer = serverService;
        _globalServerIp = globalConfig.ServerIp;
        _globalServerPort = globalConfig.ServerPort;
        _globalServerPassword = globalConfig.ServerPassword;

        LoadConfig();

        if (!backgroundMode)
        {
            LoadStudentNames();
            LoadAttendanceData();

            CheckInCommand = new RelayCommand(ExecuteCheckIn);
            CancelCheckInCommand = new RelayCommand(ExecuteCancel, p => p is StudentModel s && s.IsCheckedIn);
            ExportCommand = new RelayCommand(async _ => await ExportDataAsync());
            ImportCommand = new RelayCommand(async _ => await ImportDataAsync());
            ClearAllCommand = new RelayCommand(async _ => await ClearAllRecordsAsync());
            ShowAdminSettingsCommand = new RelayCommand(async _ => await ShowAdminSettingsAsync());
            CheckServerStatusCommand = new RelayCommand(async _ => await CheckServerStatusAsync());
            LoadFromServerCommand = new RelayCommand(async _ => await LoadFromServerAsync());
            SyncToServerCommand = new RelayCommand(async _ => await SyncToServerAsync());

            UpdateRanking();
        }

        if (Config.OnlineMode && !string.IsNullOrEmpty(_globalServerIp))
        {
            InitializeServerConnection();
        }
    }

    private void InitializeServerConnection()
    {
        _server = new ServerService();
        _server.Initialize(_globalServerIp, _globalServerPort, _globalServerPassword);
        _ = ConnectAndSyncAsync();
    }

    private async Task ConnectAndSyncAsync()
    {
        if (_server == null) return;

        try
        {
            if (Config.IsSignInTask && !string.IsNullOrEmpty(Config.SignInTaskId))
            {
                _server.TaskId = Config.SignInTaskId;
            }

            var name = Config.Name;
            if (await _server.RegisterAsync(name))
            {
                IsOnline = true;
                LastCheck = DateTime.Now.ToString("MM-dd HH:mm:ss");
                await _server.SyncConfigAsync(new ClientConfig
                {
                    School = Config.Name, Nj = "", ClassId = "",
                    Km = Config.Km, Z = Config.ButtonRows, L = Config.ButtonCols
                });
                await ApplyServerDataAsync();
            }
        }
        catch { IsOnline = false; }

        StartPeriodicSync();
    }

    public void UpdateGlobalServerConfig(string ip, int port, string password)
    {
        var wasOffline = string.IsNullOrEmpty(_globalServerIp) || _globalServerIp != ip || _globalServerPort != port;
        _globalServerIp = ip;
        _globalServerPort = port;
        _globalServerPassword = password;

        if (wasOffline && Config.OnlineMode && !string.IsNullOrEmpty(ip) && _server == null)
        {
            InitializeServerConnection();
        }
    }

    public void Dispose()
    {
        _syncCts?.Cancel();
        _syncCts?.Dispose();
    }

    private async Task ApplyServerDataAsync()
    {
        if (_server == null) return;
        var serverData = await _server.LoadDataAsync();
        if (serverData != null && serverData.Count > 0)
        {
            foreach (var (name, sa) in serverData)
            {
                var stu = Students.FirstOrDefault(s => s.Name == name);
                if (stu != null)
                {
                    stu.FirstTime = sa.FirstTime;
                    stu.Count = sa.Count;
                    stu.History = sa.History;
                }
            }
            SaveAttendanceData();
            UpdateRanking();
            OnPropertyChanged(nameof(PunchInfo));
        }
    }

    private async void StartPeriodicSync()
    {
        _syncCts?.Cancel();
        _syncCts = new CancellationTokenSource();
        var token = _syncCts.Token;
        try
        {
            while (!token.IsCancellationRequested)
            {
                await Task.Delay(5000, token);
                if (!IsOnline || _server == null) continue;

                try
                {
                    var serverData = await _server.LoadDataAsync();
                    if (serverData != null)
                    {
                        bool changed = false;
                        foreach (var (name, sa) in serverData)
                        {
                            var stu = Students.FirstOrDefault(s => s.Name == name);
                            if (stu != null && stu.FirstTime != sa.FirstTime)
                            {
                                stu.FirstTime = sa.FirstTime; stu.Count = sa.Count;
                                changed = true;
                            }
                        }
                        if (changed) { UpdateRanking(); OnPropertyChanged(nameof(PunchInfo)); }
                    }
                    var online = await _server.CheckStatusAsync();
                    if (IsOnline != online) IsOnline = online;
                    LastCheck = DateTime.Now.ToString("MM-dd HH:mm:ss");
                }
                catch { /* ignore sync errors */ }
            }
        }
        catch (OperationCanceledException) { }
    }

    // ---- 打卡操作 ----
    private void ExecuteCheckIn(object? param)
    {
        if (param is not StudentModel stu || stu.IsCheckedIn) return;
        var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        stu.FirstTime = now; stu.Count++; stu.History.Add(now);
        PunchedCount = Students.Count(s => s.IsCheckedIn);
        UpdateRanking(); SaveAttendanceData();
        StatusMessage = $"{stu.Name} 打卡成功! 时间: {now}";
        if (IsOnline && _server != null) _ = _server.SyncDataAsync(GetAttendanceDict());
    }

    private async void ExecuteCancel(object? param)
    {
        if (param is not StudentModel stu || !stu.IsCheckedIn) return;

        bool confirm = await DialogHelper.ConfirmAsync($"确定取消 {stu.Name} 的打卡？");
        if (!confirm) return;

        if (stu.History.Count > 0)
        {
            var removed = stu.History.Last();
            stu.History.RemoveAt(stu.History.Count - 1);
            if (stu.FirstTime == removed)
                stu.FirstTime = stu.History.Count > 0 ? stu.History.First() : null;
            stu.Count = Math.Max(0, stu.Count - 1);
        }
        else if (stu.FirstTime != null)
        {
            stu.FirstTime = null;
            stu.Count = 0;
        }

        PunchedCount = Students.Count(s => s.IsCheckedIn);
        UpdateRanking();
        SaveAttendanceData();
        StatusMessage = $"{stu.Name} 的打卡已取消";
        if (IsOnline && _server != null) _ = _server.SyncDataAsync(GetAttendanceDict());
    }

    public void UpdateRanking()
    {
        Ranking.Clear();
        var punched = Students.Where(s => s.IsCheckedIn).OrderBy(s => s.FirstTime).ToList();
        int rank = 1;
        foreach (var s in punched)
        {
            var t = s.FirstTime;
            var timeStr = (t != null && t.Length >= 16) ? t[11..16] : (t ?? "--:--");
            Ranking.Add(new RankingItem { Rank = rank++, Name = s.Name, Time = timeStr });
        }
    }

    // ---- 服务器操作 ----
    public async Task CheckServerStatusAsync()
    {
        try
        {
            if (string.IsNullOrEmpty(_globalServerIp))
            {
                await DialogHelper.AlertAsync("未配置服务器地址，请在设置中配置。");
                return;
            }
            if (_server == null)
            {
                _server = new ServerService();
                _server.Initialize(_globalServerIp, _globalServerPort, _globalServerPassword);
            }
            IsOnline = await _server.CheckStatusAsync();
            LastCheck = DateTime.Now.ToString("MM-dd HH:mm:ss");
            StatusMessage = IsOnline ? "服务器在线" : "服务器离线";
        }
        catch (Exception ex) { StatusMessage = $"状态检查失败: {ex.Message}"; }
    }

    public async Task LoadFromServerAsync()
    {
        if (!IsOnline) { await DialogHelper.AlertAsync("服务器不在线"); return; }
        await ApplyServerDataAsync();
        StatusMessage = "已从服务器加载数据";
    }

    public async Task SyncToServerAsync()
    {
        if (!IsOnline) { await DialogHelper.AlertAsync("服务器不在线"); return; }
        if (_server != null) await _server.SyncDataAsync(GetAttendanceDict());
        StatusMessage = "数据已同步到服务器";
    }

    // ---- 管理员设置 ----
    private async Task ShowAdminSettingsAsync()
    {
        string name = await DialogHelper.PromptAsync("任务设置", "任务名称:", Config.Name);
        if (name == null) return;
        string km = await DialogHelper.PromptAsync("任务设置", "课程名称:", Config.Km);
        if (km == null) return;
        string rowsStr = await DialogHelper.PromptAsync("任务设置", "按钮行数:", Config.ButtonRows.ToString());
        if (rowsStr == null || !int.TryParse(rowsStr, out int rows)) return;
        string colsStr = await DialogHelper.PromptAsync("任务设置", "按钮列数:", Config.ButtonCols.ToString());
        if (colsStr == null || !int.TryParse(colsStr, out int cols)) return;

        Config.Name = name; Config.Km = km;
        Config.ButtonRows = rows; Config.ButtonCols = cols;
        SaveConfig();
        StatusMessage = "任务设置已保存";
        OnPropertyChanged(nameof(TabDisplayName));
    }

    // ---- 数据文件操作 ----
    private void LoadConfig()
    {
        if (!File.Exists(_configFile)) return;
        try
        {
            var c = JsonSerializer.Deserialize<TabConfig>(File.ReadAllText(_configFile));
            if (c != null)
            {
                Config.Name = c.Name; Config.Km = c.Km;
                Config.ButtonRows = c.ButtonRows; Config.ButtonCols = c.ButtonCols;
                Config.OnlineMode = c.OnlineMode;
            }
        }
        catch { }
    }

    private void LoadStudentNames()
    {
        if (!File.Exists(_nameFile))
            File.WriteAllLines(_nameFile, Enumerable.Range(1, 40).Select(i => $"学生{i}"));
        foreach (var name in File.ReadAllLines(_nameFile).Select(l => l.Trim()).Where(l => l.Length > 0))
            Students.Add(new StudentModel { Name = name });
    }

    private void LoadAttendanceData()
    {
        if (!File.Exists(_dataFile)) return;
        foreach (var line in File.ReadAllLines(_dataFile))
        {
            var parts = line.Split(':');
            if (parts.Length < 3) continue;
            var stu = Students.FirstOrDefault(s => s.Name == parts[0]);
            if (stu != null)
            {
                stu.Count = int.TryParse(parts[1], out var c) ? c : 0;
                stu.FirstTime = string.IsNullOrEmpty(parts[2]) ? null : parts[2];
                if (parts.Length > 3 && !string.IsNullOrEmpty(parts[3]))
                    stu.History = parts[3].Split('|', StringSplitOptions.RemoveEmptyEntries).ToList();
            }
        }
        TotalStudents = Students.Count;
        PunchedCount = Students.Count(s => s.IsCheckedIn);
    }

    private void SaveAttendanceData()
    {
        File.WriteAllLines(_dataFile, Students.Select(s =>
            $"{s.Name}:{s.Count}:{s.FirstTime ?? ""}:{string.Join("|", s.History)}"));
    }

    public void SaveConfig()
    {
        File.WriteAllText(_configFile, JsonSerializer.Serialize(Config, new JsonSerializerOptions { WriteIndented = true }));
    }

    // ---- 导入/导出 ----
    private async Task ExportDataAsync()
    {
        try
        {
            var fileName = $"{Config.Name}_{Config.Km}_打卡数据.csv";
            var filePath = Path.Combine(FileSystem.CacheDirectory, fileName);
            var lines = new List<string> { "姓名,打卡时间,打卡次数,历史记录" };
            foreach (var s in Students)
                lines.Add($"{s.Name},{s.FirstTime ?? "未打卡"},{s.Count},{string.Join("|", s.History)}");
            File.WriteAllLines(filePath, lines);

            await Share.Default.RequestAsync(new ShareFileRequest
            {
                Title = "导出打卡数据",
                File = new ShareFile(filePath)
            });
            StatusMessage = $"数据已导出到: {fileName}";
        }
        catch (Exception ex)
        {
            await DialogHelper.AlertAsync($"导出失败: {ex.Message}");
        }
    }

    private async Task ImportDataAsync()
    {
        try
        {
            var result = await FilePicker.Default.PickAsync(new PickOptions
            {
                PickerTitle = "选择CSV文件",
                FileTypes = new FilePickerFileType(new Dictionary<DevicePlatform, IEnumerable<string>>
                {
                    { DevicePlatform.WinUI, new[] { ".csv" } },
                    { DevicePlatform.macOS, new[] { "public.comma-separated-values-text" } },
                })
            });

            if (result == null) return;

            using var stream = await result.OpenReadAsync();
            using var reader = new StreamReader(stream);
            var import = new Dictionary<string, StudentModel>();
            bool isFirst = true;
            while (!reader.EndOfStream)
            {
                var line = await reader.ReadLineAsync();
                if (isFirst) { isFirst = false; continue; }
                if (string.IsNullOrEmpty(line)) continue;
                var parts = line.Split(',');
                if (parts.Length < 3) continue;
                import[parts[0]] = new StudentModel
                {
                    Name = parts[0], FirstTime = parts[1] != "未打卡" ? parts[1] : null,
                    Count = int.TryParse(parts[2], out var c) ? c : 0,
                    History = parts.Length > 3 ? parts[3].Split('|', StringSplitOptions.RemoveEmptyEntries).ToList() : new()
                };
            }

            foreach (var stu in Students)
                if (import.TryGetValue(stu.Name, out var d)) { stu.FirstTime = d.FirstTime; stu.Count = d.Count; stu.History = d.History; }
            PunchedCount = Students.Count(s => s.IsCheckedIn);
            UpdateRanking(); SaveAttendanceData();
            StatusMessage = $"数据已从 {result.FileName} 导入";
        }
        catch (Exception ex)
        {
            await DialogHelper.AlertAsync($"导入失败: {ex.Message}");
        }
    }

    // ---- 清空 ----
    private async Task ClearAllRecordsAsync()
    {
        for (int i = 3; i > 0; i--)
        {
            bool c = await DialogHelper.ConfirmAsync(
                $"确定要清空「{TabDisplayName}」的所有打卡记录？\n此操作不可恢复！还有 {i} 次警告", "警告");
            if (!c) return;
        }
        bool final = await DialogHelper.ConfirmAsync("最后一次确认！确定要清空所有打卡记录？", "最终确认");
        if (!final) return;

        foreach (var stu in Students) { stu.FirstTime = null; stu.Count = 0; stu.History.Clear(); }
        PunchedCount = 0; UpdateRanking(); SaveAttendanceData();
        StatusMessage = "所有打卡记录已清空";
        if (IsOnline && _server != null) _ = _server.SyncDataAsync(GetAttendanceDict());
    }

    private Dictionary<string, StudentAttendance> GetAttendanceDict() =>
        Students.ToDictionary(s => s.Name, s => new StudentAttendance
        { Name = s.Name, Count = s.Count, FirstTime = s.FirstTime, History = s.History });

    public void NotifyPropertyChanged(string propertyName) => OnPropertyChanged(propertyName);

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
