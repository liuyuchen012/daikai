using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CheckIn.Client.Models;
using CheckIn.Client.Services;
using CheckIn.Shared.Models;
using Microsoft.Win32;

namespace CheckIn.Client.ViewModels;

/// <summary>
/// 每个标签页的独立配置，持久化到各任务目录下的 config.json
/// 包括任务名称、课程、按钮行列布局和在线模式开关
/// </summary>
public class TabConfig
{
    /// <summary>任务名称（如"三（1）班"）</summary>
    public string Name { get; set; } = "";
    /// <summary>课程名称（如"数学"）</summary>
    public string Km { get; set; } = "";
    /// <summary>学生按钮网格行数</summary>
    public int ButtonRows { get; set; } = 6;
    /// <summary>学生按钮网格列数</summary>
    public int ButtonCols { get; set; } = 6;
    /// <summary>是否启用在线模式</summary>
    public bool OnlineMode { get; set; } = true;
    /// <summary>是否为签到任务</summary>
    public bool IsSignInTask { get; set; } = false;
    /// <summary>签到任务的服务器 TaskId（如 signin_abc123）</summary>
    public string? SignInTaskId { get; set; }
}

/// <summary>
/// 标签页信息（持久化到 workspace.json，用于恢复工作区状态）
/// </summary>
public class TabInfo
{
    /// <summary>标签页唯一 ID（8 位 GUID 短码）</summary>
    public string Id { get; set; } = "";
    /// <summary>标签页显示名称</summary>
    public string Name { get; set; } = "";
}

/// <summary>
/// 每个打卡任务标签页的 ViewModel，包含独立的学生列表、排名、数据文件和服务器连接
/// 实现 IDisposable 以支持后台同步的资源释放
/// </summary>
public class TaskTabViewModel : INotifyPropertyChanged, IDisposable
{
    /// <summary>标签页唯一 ID</summary>
    public string TabId { get; }
    /// <summary>标签页配置（任务名称、课程、行列数）</summary>
    public TabConfig Config { get; } = new();

    private readonly string _tabDir;       // 标签页数据目录
    private readonly string _dataFile;     // 打卡数据文件路径
    private readonly string _nameFile;     // 学生名单文件路径
    private readonly string _configFile;   // 标签页配置文件路径
    private ServerService? _server;        // 服务器连接服务
    private CancellationTokenSource? _syncCts; // 周期性同步取消令牌

    // 全局服务器配置（由 Workspace 注入，所有标签页共享）
    private string _globalServerIp = "";
    private int _globalServerPort = 5250;
    private string _globalServerPassword = "";

    // 配置版本追踪（检测服务端推送的配置变更）
    private int _lastAppliedConfigVersion;
    private DateTime _lastConfigCheck = DateTime.MinValue;
    // 集控平台更新弹窗去重：记录已提示过的版本，避免重复弹窗
    private string _serverUpdateShownVersion = "";

    /// <summary>服务端推送了新任务配置时触发</summary>
    public event Action<List<PendingTaskConfig>>? PendingTasksReceived;

    /// <summary>集控平台自身检测到新版本时触发（用于向管理员弹窗）</summary>
    public event Action<string, string>? ServerUpdateAvailable;

    /// <summary>学生数据集合（绑定到打卡按钮网格）</summary>
    public ObservableCollection<StudentModel> Students { get; } = new();
    /// <summary>打卡排名集合（绑定到左侧排名列表）</summary>
    public ObservableCollection<RankingItem> Ranking { get; } = new();

    // ---- 在线状态相关属性 ----
    private bool _isOnline;
    /// <summary>是否与服务器保持在线连接</summary>
    public bool IsOnline
    {
        get => _isOnline;
        set { _isOnline = value; OnPropertyChanged(); OnPropertyChanged(nameof(StatusText)); OnPropertyChanged(nameof(StatusColor)); }
    }

    private string _statusMessage = "就绪";
    /// <summary>状态栏消息文本</summary>
    public string StatusMessage { get => _statusMessage; set { _statusMessage = value; OnPropertyChanged(); } }

    private string _lastCheck = "从未检查";
    /// <summary>上次服务器状态检查时间</summary>
    public string LastCheck { get => _lastCheck; set { _lastCheck = value; OnPropertyChanged(); } }

    /// <summary>服务器状态描述文本</summary>
    public string StatusText => IsOnline ? "服务器: 在线" : "服务器: 离线";
    /// <summary>服务器状态指示灯颜色</summary>
    public string StatusColor => IsOnline ? "#34a853" : "#ea4335";

    private int _totalStudents;
    /// <summary>学生总人数</summary>
    public int TotalStudents { get => _totalStudents; set { _totalStudents = value; OnPropertyChanged(); OnPropertyChanged(nameof(PunchInfo)); } }

    private int _punchedCount;
    /// <summary>已打卡人数</summary>
    public int PunchedCount { get => _punchedCount; set { _punchedCount = value; OnPropertyChanged(); OnPropertyChanged(nameof(PunchInfo)); OnPropertyChanged(nameof(PunchPercent)); } }

    /// <summary>打卡统计信息展示文本</summary>
    public string PunchInfo => $"总人数: {TotalStudents}  |  已打卡: {PunchedCount}  ({PunchPercent})";
    /// <summary>打卡百分比</summary>
    public string PunchPercent => TotalStudents > 0 ? $"{PunchedCount * 100.0 / TotalStudents:F1}%" : "0%";

    /// <summary>标签栏显示名称（格式：任务名 - 课程名）</summary>
    public string TabDisplayName => string.IsNullOrEmpty(Config.Km) ? Config.Name : $"{Config.Name} - {Config.Km}";

    // ---- 标签页级别的命令（仅非后台模式初始化，使用 null! 抑制 CS8618） ----
    /// <summary>学生打卡命令</summary>
    public ICommand CheckInCommand { get; } = null!;
    /// <summary>取消学生打卡命令</summary>
    public ICommand CancelCheckInCommand { get; } = null!;
    /// <summary>导出打卡数据命令</summary>
    public ICommand ExportCommand { get; } = null!;
    /// <summary>导入打卡数据命令</summary>
    public ICommand ImportCommand { get; } = null!;
    /// <summary>清空所有打卡记录命令</summary>
    public ICommand ClearAllCommand { get; } = null!;
    /// <summary>打开任务设置命令</summary>
    public ICommand ShowAdminSettingsCommand { get; } = null!;
    /// <summary>检查服务器状态命令</summary>
    public ICommand CheckServerStatusCommand { get; } = null!;
    /// <summary>从服务器加载数据命令</summary>
    public ICommand LoadFromServerCommand { get; } = null!;
    /// <summary>同步数据到服务器命令</summary>
    public ICommand SyncToServerCommand { get; } = null!;

    /// <summary>
    /// 创建后台运行实例（轻量级），仅建立服务器连接，不加载 UI 数据
    /// 用于未打开标签页时仍然保持与服务器的数据同步
    /// </summary>
    /// <param name="tabId">标签页 ID</param>
    /// <param name="baseDir">应用基础目录</param>
    /// <param name="globalConfig">全局配置引用</param>
    public static TaskTabViewModel CreateBackgroundInstance(string tabId, string baseDir, AppConfig globalConfig)
    {
        var tab = new TaskTabViewModel(tabId, baseDir, globalConfig, true);
        return tab;
    }

    /// <summary>
    /// 构造函数：初始化标签页目录、加载配置和学生数据，建立后台服务器连接
    /// </summary>
    /// <param name="tabId">标签页唯一 ID</param>
    /// <param name="baseDir">应用基础目录</param>
    /// <param name="globalConfig">全局配置引用（服务器地址等）</param>
    /// <param name="backgroundMode">是否为后台模式（不加载 UI 数据）</param>
    public TaskTabViewModel(string tabId, string baseDir, AppConfig globalConfig, bool backgroundMode = false)
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
        // Version is now a compile-time constant in AppConfig

        LoadConfig();

        if (!backgroundMode)
        {
            LoadStudentNames();
            LoadAttendanceData();

            CheckInCommand = new RelayCommand(ExecuteCheckIn);
            CancelCheckInCommand = new RelayCommand(ExecuteCancel, p => p is StudentModel s && s.IsCheckedIn);
            ExportCommand = new RelayCommand(_ => ExportData());
            ImportCommand = new RelayCommand(_ => ImportData());
            ClearAllCommand = new RelayCommand(_ => ClearAllRecords());
            ShowAdminSettingsCommand = new RelayCommand(_ => ShowAdminSettings());
            CheckServerStatusCommand = new RelayCommand(async _ => await CheckServerStatusAsync());
            LoadFromServerCommand = new RelayCommand(async _ => await LoadFromServerAsync());
            SyncToServerCommand = new RelayCommand(async _ => await SyncToServerAsync());

            UpdateRanking();
        }

        // 后台连接（无论是否打开标签页）
        if (Config.OnlineMode && !string.IsNullOrEmpty(_globalServerIp))
        {
            InitializeServerConnection();
        }
    }

    /// <summary>
    /// 初始化服务器连接（后台自动连接）
    /// </summary>
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
            // 签到任务使用特定的 TaskId
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

    /// <summary>
    /// 更新全局服务器配置（从 Workspace 同步到本标签页）
    /// 如果之前离线且现在有服务器配置，自动触发重连
    /// </summary>
    public void UpdateGlobalServerConfig(string ip, int port, string password, string version)
    {
        var wasOffline = string.IsNullOrEmpty(_globalServerIp) || _globalServerIp != ip || _globalServerPort != port;
        _globalServerIp = ip;
        _globalServerPort = port;
        _globalServerPassword = password;

        // 如果之前离线且现在有服务器配置，重新连接
        if (wasOffline && Config.OnlineMode && !string.IsNullOrEmpty(ip) && _server == null)
        {
            InitializeServerConnection();
        }
    }

    /// <summary>
    /// 停止后台同步并释放资源
    /// </summary>
    public void Dispose()
    {
        _syncCts?.Cancel();
        _syncCts?.Dispose();
        // _server 无需 Dispose（没有实现 IDisposable）
    }

    // ---- 在线模式：从服务器拉取数据 ----
    /// <summary>
    /// 从服务器加载最新打卡数据并应用到本地学生列表
    /// </summary>
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
                else
                {
                    // BUG FIX: 添加服务器有但本地没有的学生（新签到学生）
                    var newStu = new StudentModel { Name = name, FirstTime = sa.FirstTime, Count = sa.Count, History = sa.History };
                    Students.Add(newStu);
                }
            }
            SaveAttendanceData();
            UpdateRanking();
            OnPropertyChanged(nameof(PunchInfo));
        }
    }

    /// <summary>
    /// 启动周期性同步任务：每 5 秒从服务器拉取数据、检测配置变更并更新在线状态
    /// </summary>
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
                    // ① 拉取打卡数据
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

                    // ② 检测服务端配置推送（每 15 秒检查一次，避免频繁请求）
                    var now = DateTime.Now;
                    if ((now - _lastConfigCheck).TotalSeconds >= 15)
                    {
                        _lastConfigCheck = now;
                        var remoteConfig = await _server.LoadConfigAsync();
                        if (remoteConfig != null && remoteConfig.ConfigVersion > _lastAppliedConfigVersion)
                        {
                            _lastAppliedConfigVersion = remoteConfig.ConfigVersion;

                            // 应用设备名称变更
                            if (!string.IsNullOrEmpty(remoteConfig.DeviceName) && Config.Name != remoteConfig.DeviceName)
                            {
                                Config.Name = remoteConfig.DeviceName;
                                SaveConfig();
                                OnPropertyChanged(nameof(Config));
                                OnPropertyChanged(nameof(TabDisplayName));
                            }

                            // 处理待推送的任务
                            if (remoteConfig.PendingTasks is { Count: > 0 })
                            {
                                var pendingTasks = remoteConfig.PendingTasks.ToList();
                                // 通知 MainViewModel 创建新标签页
                                PendingTasksReceived?.Invoke(pendingTasks);

                                // 确认已应用的任务
                                var appliedIds = pendingTasks.Select(t => t.TaskId).ToList();
                                await _server.ConfigAppliedAsync(appliedIds);
                            }
                        }

                        // ③ 检查集控平台自身是否有新版本（每 15 秒轮询一次）
                        var upd = await _server.GetServerUpdateAsync();
                        if (upd != null && upd.Value.hasUpdate && upd.Value.latestVersion != _serverUpdateShownVersion)
                        {
                            _serverUpdateShownVersion = upd.Value.latestVersion;
                            ServerUpdateAvailable?.Invoke(upd.Value.latestVersion, upd.Value.downloadUrl);
                        }
                    }

                    // ③ 检查在线状态
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
    /// <summary>
    /// 执行学生打卡：记录首次打卡时间、累计次数、更新时间戳列表
    /// </summary>
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

    /// <summary>
    /// 取消学生打卡：移除最后一次记录，恢复首次时间或清空状态
    /// </summary>
    private void ExecuteCancel(object? param)
    {
        if (param is not StudentModel stu || !stu.IsCheckedIn) return;
        if (!ModernDialog.Confirm($"确定取消 {stu.Name} 的打卡？")) return;

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

    /// <summary>
    /// 刷新打卡排名：按首次打卡时间升序排列已打卡学生
    /// </summary>
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
    /// <summary>
    /// 手动检查服务器在线状态
    /// </summary>
    public async Task CheckServerStatusAsync()
    {
        try
        {
            if (string.IsNullOrEmpty(_globalServerIp))
            {
                ModernDialog.Alert("未配置服务器地址，请在 [管理员设置] 中配置。");
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

    /// <summary>
    /// 手动从服务器加载最新打卡数据
    /// </summary>
    public async Task LoadFromServerAsync()
    {
        if (!IsOnline) { ModernDialog.Alert("服务器不在线"); return; }
        await ApplyServerDataAsync();
        StatusMessage = "已从服务器加载数据";
    }

    /// <summary>
    /// 手动将本地数据同步到服务器
    /// </summary>
    public async Task SyncToServerAsync()
    {
        if (!IsOnline) { ModernDialog.Alert("服务器不在线"); return; }
        if (_server != null) await _server.SyncDataAsync(GetAttendanceDict());
        StatusMessage = "数据已同步到服务器";
    }

    // ---- 管理员设置（标签页级别） ----
    /// <summary>
    /// 显示标签页级别的任务设置对话框
    /// </summary>
    private void ShowAdminSettings()
    {
        var vm = new TabAdminSettingsVm
        {
            Name = Config.Name, Km = Config.Km,
            ButtonRows = Config.ButtonRows, ButtonCols = Config.ButtonCols,
        };
        vm.SaveAction = () =>
        {
            Config.Name = vm.Name; Config.Km = vm.Km;
            Config.ButtonRows = vm.ButtonRows; Config.ButtonCols = vm.ButtonCols;
            SaveConfig();
            StatusMessage = "任务设置已保存";
            OnPropertyChanged(nameof(TabDisplayName));
        };
        ShowDialog("任务设置", vm, CreateAdminSettingsView);
    }

    /// <summary>
    /// 通用的无边框对话框创建方法
    /// </summary>
    private static void ShowDialog<T>(string title, T? vm, Func<T?, FrameworkElement> createView, double width = 420, double height = 320)
    {
        var win = new Window
        {
            Title = title, Width = width, Height = height,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false,
            Content = createView(vm)
        };
        if (vm is IDialogVm dvm) dvm.CloseAction = win.Close;
        win.ShowDialog();
    }

    /// <summary>
    /// 构建标签页管理员设置对话框的 UI 视图
    /// </summary>
    private FrameworkElement CreateAdminSettingsView(TabAdminSettingsVm? vm)
    {
        if (vm == null) return new StackPanel { Margin = new Thickness(15) };
        var panel = new StackPanel { Margin = new Thickness(15) };
        AddFieldRow(panel, "任务名称:", vm.Name, v => vm.Name = v);
        AddFieldRow(panel, "课程名称:", vm.Km, v => vm.Km = v);
        AddFieldRow(panel, "按钮行数:", vm.ButtonRows.ToString(), v => { if (int.TryParse(v, out var n)) vm.ButtonRows = n; });
        AddFieldRow(panel, "按钮列数:", vm.ButtonCols.ToString(), v => { if (int.TryParse(v, out var n)) vm.ButtonCols = n; });
        AddSaveCancel(panel, vm);
        return panel;
    }

    /// <summary>
    /// 向面板添加一行表单字段（标签 + 输入框）
    /// </summary>
    private static void AddFieldRow(System.Windows.Controls.Panel parent, string label, string value, Action<string> setter, bool isPassword = false)
    {
        var row = new System.Windows.Controls.WrapPanel { Margin = new Thickness(4) };
        row.Children.Add(new System.Windows.Controls.TextBlock { Text = label, Width = 90, VerticalAlignment = VerticalAlignment.Center });
        if (isPassword)
        {
            var tb = new System.Windows.Controls.PasswordBox { Width = 250 };
            tb.PasswordChanged += (_, _) => setter(tb.Password);
            row.Children.Add(tb);
        }
        else
        {
            var tb = new System.Windows.Controls.TextBox { Text = value, Width = 250 };
            tb.TextChanged += (_, _) => setter(tb.Text);
            row.Children.Add(tb);
        }
        parent.Children.Add(row);
    }

    /// <summary>
    /// 向面板添加保存/取消按钮行
    /// </summary>
    private static void AddSaveCancel(System.Windows.Controls.Panel parent, IDialogVm vm)
    {
        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Center,
            Margin = new Thickness(0, 15, 0, 0)
        };
        var save = new System.Windows.Controls.Button { Content = "保存", Width = 80, Height = 30, Margin = new Thickness(5) };
        save.Click += (_, _) => vm.SaveAction?.Invoke();
        var cancel = new System.Windows.Controls.Button { Content = "取消", Width = 80, Height = 30, Margin = new Thickness(5) };
        cancel.Click += (_, _) => vm.CloseAction?.Invoke();
        btnPanel.Children.Add(save); btnPanel.Children.Add(cancel);
        parent.Children.Add(btnPanel);
    }

    // ---- 数据文件操作 ----
    /// <summary>
    /// 从标签页目录加载任务配置（config.json）
    /// </summary>
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

    /// <summary>
    /// 加载学生名单（name.txt），如果文件不存在则生成 40 名默认学生
    /// </summary>
    public void LoadStudentNames()
    {
        if (!File.Exists(_nameFile))
            File.WriteAllLines(_nameFile, Enumerable.Range(1, 40).Select(i => $"学生{i}"));
        Students.Clear(); // BUG FIX: 清空已有的学生列表，避免多次调用导致追加而非替换
        foreach (var name in File.ReadAllLines(_nameFile).Select(l => l.Trim()).Where(l => l.Length > 0))
            Students.Add(new StudentModel { Name = name });
    }

    /// <summary>
    /// 从 attendance.dat 文件加载历史打卡数据，格式为 姓名:次数:首次时间:历史记录|...
    /// </summary>
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

    /// <summary>
    /// 保存打卡数据到磁盘（attendance.dat），每行格式：姓名:次数:首次时间:历史|记录|列表
    /// </summary>
    private void SaveAttendanceData()
    {
        File.WriteAllLines(_dataFile, Students.Select(s =>
            $"{s.Name}:{s.Count}:{s.FirstTime ?? ""}:{string.Join("|", s.History)}"));
    }

    /// <summary>
    /// 保存标签页配置到 config.json
    /// </summary>
    public void SaveConfig()
    {
        File.WriteAllText(_configFile, JsonSerializer.Serialize(Config, new JsonSerializerOptions { WriteIndented = true }));
    }

    // ---- 导入/导出 ----
    /// <summary>
    /// 导出打卡数据为 CSV 文件（姓名,打卡时间,打卡次数,历史记录）
    /// </summary>
    private void ExportData()
    {
        var dlg = new SaveFileDialog
        {
            Filter = "CSV文件|*.csv", DefaultExt = ".csv",
            FileName = $"{Config.Name}_{Config.Km}_打卡数据.csv"
        };
        if (dlg.ShowDialog() != true) return;
        var lines = new List<string> { "姓名,打卡时间,打卡次数,历史记录" };
        foreach (var s in Students)
            lines.Add($"{s.Name},{s.FirstTime ?? "未打卡"},{s.Count},{string.Join("|", s.History)}");
        File.WriteAllLines(dlg.FileName, lines);
        StatusMessage = $"数据已导出到: {dlg.FileName}";
    }

    /// <summary>
    /// 从 CSV 文件导入打卡数据，覆盖当前学生数据
    /// </summary>
    private void ImportData()
    {
        var dlg = new OpenFileDialog { Filter = "CSV文件|*.csv" };
        if (dlg.ShowDialog() != true) return;
        var import = new Dictionary<string, StudentModel>();
        foreach (var line in File.ReadAllLines(dlg.FileName).Skip(1))
        {
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
        StatusMessage = $"数据已从 {dlg.FileName} 导入";
    }

    // ---- 清空 ----
    /// <summary>
    /// 清空所有打卡记录，需要多次弹窗确认防止误操作
    /// </summary>
    private void ClearAllRecords()
    {
        for (int i = 3; i > 0; i--)
            if (!ModernDialog.Confirm($"确定要清空「{TabDisplayName}」的所有打卡记录？\n此操作不可恢复！还有 {i} 次警告", "警告"))
                return;
        if (!ModernDialog.Confirm("最后一次确认！确定要清空所有打卡记录？", "最终确认")) return;

        foreach (var stu in Students) { stu.FirstTime = null; stu.Count = 0; stu.History.Clear(); }
        PunchedCount = 0; UpdateRanking(); SaveAttendanceData();
        StatusMessage = "所有打卡记录已清空";
        if (IsOnline && _server != null) _ = _server.SyncDataAsync(GetAttendanceDict());
    }

    /// <summary>
    /// 将当前学生列表转换为服务器同步所需的字典格式
    /// </summary>
    private Dictionary<string, StudentAttendance> GetAttendanceDict() =>
        Students.ToDictionary(s => s.Name, s => new StudentAttendance
        { Name = s.Name, Count = s.Count, FirstTime = s.FirstTime, History = s.History });

    /// <summary>
    /// 手动触发属性变更通知（供外部调用）
    /// </summary>
    public void NotifyPropertyChanged(string propertyName) => OnPropertyChanged(propertyName);

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

// ---- 标签页管理员设置 VM ----
/// <summary>
/// 标签页级别的管理员设置视图模型
/// </summary>
public class TabAdminSettingsVm : IDialogVm
{
    /// <summary>任务名称</summary>
    public string Name { get; set; } = "";
    /// <summary>课程名称</summary>
    public string Km { get; set; } = "";
    /// <summary>按钮行数</summary>
    public int ButtonRows { get; set; } = 6;
    /// <summary>按钮列数</summary>
    public int ButtonCols { get; set; } = 6;
    /// <summary>关闭对话框的回调</summary>
    public Action? CloseAction { get; set; }
    /// <summary>保存设置的回调</summary>
    public Action? SaveAction { get; set; }
}
