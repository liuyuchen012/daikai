using System.Collections.ObjectModel;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Mobile.Models;
using CheckIn.Client.Mobile.Services;
using CheckIn.Shared.Models;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// Tab-level configuration persisted per task.
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
/// Tab info for workspace persistence.
/// </summary>
public class TabInfo
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
}

/// <summary>
/// Ranking display item.
/// </summary>
public class RankingItem
{
    public int Rank { get; set; }
    public string Name { get; set; } = "";
    public string Time { get; set; } = "";
}

/// <summary>
/// ViewModel for each check-in task tab. Manages student list, rankings,
/// server connection, and periodic sync. Adapted from WPF version.
/// </summary>
public class TaskTabViewModel : BaseViewModel, IDisposable
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

    public ObservableCollection<StudentModel> Students { get; } = new();
    public ObservableCollection<RankingItem> Ranking { get; } = new();

    private bool _isOnline;
    public bool IsOnline
    {
        get => _isOnline;
        set { SetProperty(ref _isOnline, value); OnPropertyChanged(nameof(StatusText)); OnPropertyChanged(nameof(StatusColor)); }
    }

    private string _statusMessage = "Ready";
    public string StatusMessage
    {
        get => _statusMessage;
        set => SetProperty(ref _statusMessage, value);
    }

    private string _lastCheck = "Never";
    public string LastCheck
    {
        get => _lastCheck;
        set => SetProperty(ref _lastCheck, value);
    }

    public string StatusText => IsOnline ? "Server: Online" : "Server: Offline";
    public Color StatusColor => IsOnline ? Color.FromArgb("#34a853") : Color.FromArgb("#ea4335");

    private int _totalStudents;
    public int TotalStudents
    {
        get => _totalStudents;
        set { SetProperty(ref _totalStudents, value); OnPropertyChanged(nameof(PunchInfo)); }
    }

    private int _punchedCount;
    public int PunchedCount
    {
        get => _punchedCount;
        set { SetProperty(ref _punchedCount, value); OnPropertyChanged(nameof(PunchInfo)); OnPropertyChanged(nameof(PunchPercent)); }
    }

    public string PunchInfo => $"Total: {TotalStudents}  |  Checked: {PunchedCount}  ({PunchPercent})";
    public string PunchPercent => TotalStudents > 0 ? $"{PunchedCount * 100.0 / TotalStudents:F1}%" : "0%";

    public string TabDisplayName =>
        string.IsNullOrEmpty(Config.Km) ? Config.Name : $"{Config.Name} - {Config.Km}";

    // Commands
    public ICommand CheckInCommand { get; }
    public ICommand CancelCheckInCommand { get; }
    public ICommand ClearAllCommand { get; }
    public ICommand SyncToServerCommand { get; }
    public ICommand LoadFromServerCommand { get; }
    public ICommand CheckServerStatusCommand { get; }

    /// <summary>
    /// Constructor: load config, student data, and initialize server connection.
    /// </summary>
    public TaskTabViewModel(string tabId, string baseDir, AppConfig globalConfig)
    {
        TabId = tabId;
        _tabDir = Path.Combine(baseDir, "data", "tabs", tabId);
        Directory.CreateDirectory(_tabDir);
        _dataFile = Path.Combine(_tabDir, "attendance.dat");
        _nameFile = Path.Combine(_tabDir, "name.txt");
        _configFile = Path.Combine(_tabDir, "config.json");

        _globalServerIp = globalConfig.ServerIp;
        _globalServerPort = globalConfig.ServerPort;
        _globalServerPassword = globalConfig.ServerPassword;

        Title = TabDisplayName;

        LoadConfig();
        LoadStudentNames();
        LoadAttendanceData();

        CheckInCommand = new RelayCommand(ExecuteCheckIn);
        CancelCheckInCommand = new RelayCommand(ExecuteCancel, p => p is StudentModel s && s.IsCheckedIn);
        ClearAllCommand = new RelayCommand(ExecuteClearAll);
        SyncToServerCommand = new RelayCommand(async _ => await SyncToServerAsync());
        LoadFromServerCommand = new RelayCommand(async _ => await LoadFromServerAsync());
        CheckServerStatusCommand = new RelayCommand(async _ => await CheckServerStatusAsync());

        UpdateRanking();

        if (Config.OnlineMode && !string.IsNullOrEmpty(_globalServerIp))
        {
            InitializeServerConnection();
        }
    }

    /// <summary>
    /// Background-mode constructor (server connection only, no UI data).
    /// </summary>
    private TaskTabViewModel(string tabId, string baseDir, AppConfig globalConfig, bool backgroundMode)
    {
        TabId = tabId;
        _tabDir = Path.Combine(baseDir, "data", "tabs", tabId);
        _dataFile = Path.Combine(_tabDir, "attendance.dat");
        _nameFile = Path.Combine(_tabDir, "name.txt");
        _configFile = Path.Combine(_tabDir, "config.json");

        _globalServerIp = globalConfig.ServerIp;
        _globalServerPort = globalConfig.ServerPort;
        _globalServerPassword = globalConfig.ServerPassword;

        LoadConfig();

        CheckInCommand = new RelayCommand(_ => { });
        CancelCheckInCommand = new RelayCommand(_ => { });
        ClearAllCommand = new RelayCommand(_ => { });
        SyncToServerCommand = new RelayCommand(_ => { });
        LoadFromServerCommand = new RelayCommand(_ => { });
        CheckServerStatusCommand = new RelayCommand(_ => { });

        if (Config.OnlineMode && !string.IsNullOrEmpty(_globalServerIp))
        {
            InitializeServerConnection();
        }
    }

    public static TaskTabViewModel CreateBackgroundInstance(string tabId, string baseDir, AppConfig globalConfig)
        => new(tabId, baseDir, globalConfig, true);

    // ---- Server Connection ----
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
                _server.TaskId = Config.SignInTaskId;

            if (await _server.RegisterAsync(Config.Name))
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

    public void UpdateGlobalServerConfig(string ip, int port, string password, string version)
    {
        var wasOffline = string.IsNullOrEmpty(_globalServerIp) || _globalServerIp != ip || _globalServerPort != port;
        _globalServerIp = ip;
        _globalServerPort = port;
        _globalServerPassword = password;

        if (wasOffline && Config.OnlineMode && !string.IsNullOrEmpty(ip) && _server == null)
            InitializeServerConnection();
    }

    public void Dispose()
    {
        _syncCts?.Cancel();
        _syncCts?.Dispose();
    }

    // ---- Data Sync ----
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
                        if (changed) { UpdateRanking(); }
                    }
                    var online = await _server.CheckStatusAsync();
                    IsOnline = online;
                    LastCheck = DateTime.Now.ToString("MM-dd HH:mm:ss");
                }
                catch { }
            }
        }
        catch (OperationCanceledException) { }
    }

    // ---- Check-in Operations ----
    private void ExecuteCheckIn(object? param)
    {
        if (param is not StudentModel stu || stu.IsCheckedIn) return;
        var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        stu.FirstTime = now;
        stu.Count++;
        stu.History.Add(now);
        PunchedCount = Students.Count(s => s.IsCheckedIn);
        UpdateRanking();
        SaveAttendanceData();
        StatusMessage = $"{stu.Name} checked in! Time: {now}";
        if (IsOnline && _server != null) _ = _server.SyncDataAsync(GetAttendanceDict());
    }

    private void ExecuteCancel(object? param)
    {
        if (param is not StudentModel stu || !stu.IsCheckedIn) return;

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
        StatusMessage = $"{stu.Name} check-in cancelled";
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

    // ---- Server Operations ----
    public async Task CheckServerStatusAsync()
    {
        try
        {
            if (string.IsNullOrEmpty(_globalServerIp))
            {
                StatusMessage = "No server configured";
                return;
            }
            _server ??= new ServerService();
            _server.Initialize(_globalServerIp, _globalServerPort, _globalServerPassword);
            IsOnline = await _server.CheckStatusAsync();
            LastCheck = DateTime.Now.ToString("MM-dd HH:mm:ss");
            StatusMessage = IsOnline ? "Server online" : "Server offline";
        }
        catch (Exception ex) { StatusMessage = $"Status check failed: {ex.Message}"; }
    }

    public async Task LoadFromServerAsync()
    {
        if (!IsOnline) { StatusMessage = "Server offline"; return; }
        await ApplyServerDataAsync();
        StatusMessage = "Data loaded from server";
    }

    public async Task SyncToServerAsync()
    {
        if (!IsOnline) { StatusMessage = "Server offline"; return; }
        if (_server != null) await _server.SyncDataAsync(GetAttendanceDict());
        StatusMessage = "Data synced to server";
    }

    private void ExecuteClearAll(object? _)
    {
        foreach (var stu in Students)
        {
            stu.FirstTime = null;
            stu.Count = 0;
            stu.History.Clear();
        }
        PunchedCount = 0;
        UpdateRanking();
        SaveAttendanceData();
        StatusMessage = "All records cleared";
        if (IsOnline && _server != null) _ = _server.SyncDataAsync(GetAttendanceDict());
    }

    // ---- File Operations ----
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
                Config.IsSignInTask = c.IsSignInTask;
                Config.SignInTaskId = c.SignInTaskId;
            }
        }
        catch { }
    }

    private void LoadStudentNames()
    {
        if (!File.Exists(_nameFile))
            File.WriteAllLines(_nameFile, Enumerable.Range(1, 40).Select(i => $"Student{i}"));
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

    private Dictionary<string, StudentAttendance> GetAttendanceDict() =>
        Students.ToDictionary(s => s.Name, s => new StudentAttendance
        { Name = s.Name, Count = s.Count, FirstTime = s.FirstTime, History = s.History });
}
