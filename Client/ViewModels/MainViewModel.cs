using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using CheckIn.Client.Models;
using CheckIn.Client.Services;
using CheckIn.Shared.Models;
using Microsoft.Win32;

namespace CheckIn.Client.ViewModels;

public class MainViewModel : INotifyPropertyChanged
{
    public AppConfig Config { get; } = new();
    private readonly ServerService _server = new();

    public ObservableCollection<StudentModel> Students { get; } = new();
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

    public string WindowTitle => $"{Config.School}{Config.Nj}年{Config.ClassId}班{Config.Km}打卡 {Config.Version}  作者: 刘宇晨";

    // Commands
    public ICommand CheckInCommand { get; }
    public ICommand CancelCheckInCommand { get; }
    public ICommand ExportCommand { get; }
    public ICommand ImportCommand { get; }
    public ICommand ClearAllCommand { get; }
    public ICommand ExitCommand { get; }
    public ICommand ShowRemoteSettingsCommand { get; }
    public ICommand CheckServerStatusCommand { get; }
    public ICommand LoadFromServerCommand { get; }
    public ICommand SyncToServerCommand { get; }
    public ICommand ShowAdminSettingsCommand { get; }
    public ICommand OpenGithubCommand { get; }
    public ICommand CheckVersionCommand { get; }
    public ICommand ShowAboutCommand { get; }

    public MainViewModel()
    {
        LoadConfig();
        LoadStudentNames();
        LoadAttendanceData();

        CheckInCommand = new RelayCommand(ExecuteCheckIn);
        CancelCheckInCommand = new RelayCommand(ExecuteCancel, p => p is StudentModel s && s.IsCheckedIn);
        ExportCommand = new RelayCommand(_ => ExportData());
        ImportCommand = new RelayCommand(_ => ImportData());
        ClearAllCommand = new RelayCommand(_ => ClearAllRecords());
        ExitCommand = new RelayCommand(_ => Application.Current.Shutdown());
        ShowRemoteSettingsCommand = new RelayCommand(_ => ShowRemoteSettings());
        CheckServerStatusCommand = new RelayCommand(async _ => await CheckServerStatusAsync());
        LoadFromServerCommand = new RelayCommand(async _ => await LoadFromServerAsync());
        SyncToServerCommand = new RelayCommand(async _ => await SyncToServerAsync());
        ShowAdminSettingsCommand = new RelayCommand(_ => ShowAdminSettings());
        OpenGithubCommand = new RelayCommand(_ => OpenUrl("https://github.com/liuyuchen012/daikai"));
        CheckVersionCommand = new RelayCommand(_ => OpenUrl("https://github.com/liuyuchen012/daikai/releases"));
        ShowAboutCommand = new RelayCommand(_ => ShowAbout());

        UpdateRanking();

        if (Config.OnlineMode && !string.IsNullOrEmpty(Config.ServerIp))
        {
            _server.Initialize(Config.ServerIp, Config.ServerPort, Config.ServerPassword);
            _ = InitializeOnlineAsync();
        }
    }

    // ---- 在线模式初始化 ----
    private async Task InitializeOnlineAsync()
    {
        var name = $"{Config.School}{Config.Nj}年{Config.ClassId}班";
        if (await _server.RegisterAsync(name))
        {
            IsOnline = true;
            LastCheck = DateTime.Now.ToString("MM-dd HH:mm:ss");
            _ = _server.SyncConfigAsync(new ClientConfig
            {
                School = Config.School, Nj = Config.Nj, ClassId = Config.ClassId,
                Km = Config.Km, Z = Config.ButtonRows, L = Config.ButtonCols
            });
            await ApplyServerDataAsync();
        }
        StartPeriodicSync();
    }

    private async Task ApplyServerDataAsync()
    {
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
        while (true)
        {
            await Task.Delay(5000);
            if (!IsOnline) continue;

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

    // ---- 打卡操作 ----
    private void ExecuteCheckIn(object? param)
    {
        if (param is not StudentModel stu || stu.IsCheckedIn) return;
        var now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        stu.FirstTime = now; stu.Count++; stu.History.Add(now);
        PunchedCount = Students.Count(s => s.IsCheckedIn);
        UpdateRanking(); SaveAttendanceData();
        StatusMessage = $"{stu.Name} 打卡成功! 时间: {now}";
        if (IsOnline) _ = _server.SyncDataAsync(GetAttendanceDict());
    }

    private void ExecuteCancel(object? param)
    {
        if (param is not StudentModel stu || !stu.IsCheckedIn) return;
        if (!ModernDialog.Confirm($"确定取消 {stu.Name} 的打卡？")) return;

        // 从历史记录中移除最后一次打卡
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
            // 没有历史记录但有打卡时间 → 直接清空
            stu.FirstTime = null;
            stu.Count = 0;
        }

        // 强制刷新绑定状态
        PunchedCount = Students.Count(s => s.IsCheckedIn);
        UpdateRanking();
        SaveAttendanceData();
        StatusMessage = $"{stu.Name} 的打卡已取消";
        if (IsOnline) _ = _server.SyncDataAsync(GetAttendanceDict());
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
    private async Task CheckServerStatusAsync()
    {
        try
        {
            if (string.IsNullOrEmpty(Config.ServerIp))
            {
                ModernDialog.Alert("未配置服务器地址，请在 [远程]→[远程服务器设置] 中配置。");
                return;
            }
            _server.Initialize(Config.ServerIp, Config.ServerPort, Config.ServerPassword);
            IsOnline = await _server.CheckStatusAsync();
            LastCheck = DateTime.Now.ToString("MM-dd HH:mm:ss");
            StatusMessage = IsOnline ? "服务器在线" : "服务器离线";
        }
        catch (Exception ex)
        {
            StatusMessage = $"状态检查失败: {ex.Message}";
        }
    }

    private async Task LoadFromServerAsync()
    {
        if (!IsOnline) { ModernDialog.Alert("服务器不在线"); return; }
        await ApplyServerDataAsync();
        StatusMessage = "已从服务器加载数据";
    }

    private async Task SyncToServerAsync()
    {
        if (!IsOnline) { ModernDialog.Alert("服务器不在线"); return; }
        await _server.SyncDataAsync(GetAttendanceDict());
        StatusMessage = "数据已同步到服务器";
    }

    // ---- 远程设置 ----
    private void ShowRemoteSettings()
    {
        var vm = new RemoteSettingsVm
        {
            Ip = Config.ServerIp, Port = Config.ServerPort, Password = Config.ServerPassword
        };
        vm.SaveAction = () =>
        {
            Config.ServerIp = vm.Ip; Config.ServerPort = vm.Port; Config.ServerPassword = vm.Password;
            SaveConfig();
            _server.Initialize(vm.Ip, vm.Port, vm.Password);
            _ = InitializeOnlineAsync();
            StatusMessage = "远程设置已保存";
        };
        ShowDialog("远程服务器设置", vm, CreateRemoteSettingsView);
    }

    // ---- 管理员设置 ----
    private void ShowAdminSettings()
    {
        if (!VerifyAdminPwd("访问管理员设置")) return;
        var vm = new AdminSettingsVm
        {
            School = Config.School, Nj = Config.Nj, ClassId = Config.ClassId, Km = Config.Km,
            ButtonRows = Config.ButtonRows, ButtonCols = Config.ButtonCols,
            NewPassword = "", ConfirmPassword = ""
        };
        vm.SaveAction = () =>
        {
            Config.School = vm.School; Config.Nj = vm.Nj; Config.ClassId = vm.ClassId; Config.Km = vm.Km;
            Config.ButtonRows = vm.ButtonRows; Config.ButtonCols = vm.ButtonCols;
            if (!string.IsNullOrEmpty(vm.NewPassword))
            {
                if (vm.NewPassword != vm.ConfirmPassword)
                {
                    ModernDialog.Alert("两次输入的密码不一致", "错误");
                    return;
                }
                Config.AdminPasswordHash = HashPassword(vm.NewPassword);
            }
            SaveConfig();
            StatusMessage = "管理员设置已保存，请重启程序生效";
            ModernDialog.Alert("所有设置已保存并生效，请重启程序");
            Application.Current.Shutdown();
        };
        ShowDialog("管理员设置", vm, CreateAdminSettingsView);
    }

    // ---- 关于 ----
    private void ShowAbout()
    {
        var aboutWin = new Window
        {
            Title = "关于", Width = 360, Height = 240,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false,
            Content = new System.Windows.Controls.StackPanel
            {
                Margin = new Thickness(10),
                Children =
                {
                    new System.Windows.Controls.TextBlock { Text = "打卡系统", FontSize = 20, FontWeight = FontWeights.Bold, Margin = new Thickness(0,0,0,10) },
                    new System.Windows.Controls.TextBlock { Text = $"版本: {Config.Version}", FontSize = 14, Margin = new Thickness(0,0,0,5) },
                    new System.Windows.Controls.TextBlock { Text = "开发者: 刘宇晨", FontSize = 14, Margin = new Thickness(0,0,0,5) },
                    new System.Windows.Controls.TextBlock { Text = "联系邮箱: liuyuchen032901@outlook.com", FontSize = 14, Margin = new Thickness(0,0,0,5) },
                    new System.Windows.Controls.TextBlock { Text = "© 2026 保留全部权利", FontSize = 12, Foreground = System.Windows.Media.Brushes.Gray, Margin = new Thickness(0,20,0,0) }
                }
            }
        };
        aboutWin.ShowDialog();
    }

    // ---- 辅助方法 ----
    private static void OpenUrl(string url)
    {
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
        catch { }
    }

    private bool VerifyAdminPwd(string action)
    {
        if (string.IsNullOrEmpty(Config.AdminPasswordHash)) return true;
        var pwd = AskPassword($"需要管理员权限执行: {action}");
        return pwd != null && HashPassword(pwd) == Config.AdminPasswordHash;
    }

    private static string HashPassword(string pwd)
    {
        using var sha = System.Security.Cryptography.SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pwd)));
    }

    private string? AskPassword(string title)
    {
        string? result = null;
        var win = new Window
        {
            Title = title, Width = 320, Height = 180,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize,
            ShowInTaskbar = false
        };
        var panel = new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };
        panel.Children.Add(new System.Windows.Controls.TextBlock { Text = "请输入管理员密码:", FontSize = 14, Margin = new Thickness(0, 0, 0, 10) });
        var pwdBox = new System.Windows.Controls.PasswordBox { FontSize = 14 };
        panel.Children.Add(pwdBox);
        var btnPanel = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment = System.Windows.HorizontalAlignment.Center, Margin = new Thickness(0, 15, 0, 0) };
        var okBtn = new System.Windows.Controls.Button { Content = "确定", Width = 80, Height = 30, Margin = new Thickness(5) };
        okBtn.Click += (_, _) => { result = pwdBox.Password; win.Close(); };
        var cancelBtn = new System.Windows.Controls.Button { Content = "取消", Width = 80, Height = 30, Margin = new Thickness(5) };
        cancelBtn.Click += (_, _) => win.Close();
        btnPanel.Children.Add(okBtn); btnPanel.Children.Add(cancelBtn);
        panel.Children.Add(btnPanel);
        win.Content = panel;
        win.ShowDialog();
        return result;
    }

    private static void ShowDialog<T>(string title, T? vm, Func<T?, System.Windows.FrameworkElement> createView, double width = 420, double height = 320)
    {
        var win = new Window
        {
            Title = title, Width = width, Height = height,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize,
            ShowInTaskbar = false,
            Content = createView(vm)
        };

        if (vm is IDialogVm dvm) dvm.CloseAction = win.Close;
        win.ShowDialog();
    }

    private System.Windows.FrameworkElement CreateRemoteSettingsView(RemoteSettingsVm? vm)
    {
        if (vm == null) return EmptyPanel();
        var panel = new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };
        AddFieldRow(panel, "服务器地址:", vm.Ip, v => vm.Ip = v);
        AddFieldRow(panel, "服务器端口:", vm.Port.ToString(), v => { if (int.TryParse(v, out var n)) vm.Port = n; });
        AddFieldRow(panel, "服务器密码:", vm.Password, v => vm.Password = v, true);
        AddSaveCancel(panel, vm);
        return panel;
    }

    private System.Windows.FrameworkElement CreateAdminSettingsView(AdminSettingsVm? vm)
    {
        if (vm == null) return EmptyPanel();
        var panel = new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };
        AddFieldRow(panel, "学校名称:", vm.School, v => vm.School = v);
        AddFieldRow(panel, "年级:", vm.Nj, v => vm.Nj = v);
        AddFieldRow(panel, "班级:", vm.ClassId, v => vm.ClassId = v);
        AddFieldRow(panel, "课程名称:", vm.Km, v => vm.Km = v);
        AddFieldRow(panel, "按钮行数:", vm.ButtonRows.ToString(), v => { if (int.TryParse(v, out var n)) vm.ButtonRows = n; });
        AddFieldRow(panel, "按钮列数:", vm.ButtonCols.ToString(), v => { if (int.TryParse(v, out var n)) vm.ButtonCols = n; });
        AddFieldRow(panel, "新密码(留空不改):", vm.NewPassword, v => vm.NewPassword = v, true);
        AddFieldRow(panel, "确认密码:", vm.ConfirmPassword, v => vm.ConfirmPassword = v, true);
        AddSaveCancel(panel, vm);
        return panel;
    }

    private static System.Windows.Controls.StackPanel EmptyPanel() => new() { Margin = new Thickness(15) };

    private static void AddFieldRow(System.Windows.Controls.Panel parent, string label, string value, Action<string> setter, bool isPassword = false)
    {
        var row = new System.Windows.Controls.WrapPanel { Margin = new Thickness(4) };
        row.Children.Add(new System.Windows.Controls.TextBlock { Text = label, Width = 90, VerticalAlignment = System.Windows.VerticalAlignment.Center });
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

    private static void AddSaveCancel(System.Windows.Controls.Panel parent, IDialogVm vm)
    {
        var btnPanel = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
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
    private readonly string _dataFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "attendance.dat");
    private readonly string _nameFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "name.txt");
    private readonly string _configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");

    private void LoadConfig()
    {
        if (!File.Exists(_configFile)) return;
        try
        {
            var c = JsonSerializer.Deserialize<AppConfig>(File.ReadAllText(_configFile));
            if (c != null)
            {
                Config.School = c.School; Config.Nj = c.Nj; Config.ClassId = c.ClassId; Config.Km = c.Km;
                Config.ButtonRows = c.ButtonRows; Config.ButtonCols = c.ButtonCols;
                Config.OnlineMode = c.OnlineMode; Config.ServerIp = c.ServerIp;
                Config.ServerPort = c.ServerPort; Config.ServerPassword = c.ServerPassword;
                Config.AdminPasswordHash = c.AdminPasswordHash;
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
    private void ExportData()
    {
        var dlg = new SaveFileDialog { Filter = "CSV文件|*.csv", DefaultExt = ".csv" };
        if (dlg.ShowDialog() != true) return;
        var lines = new List<string> { "姓名,打卡时间,打卡次数,历史记录" };
        foreach (var s in Students)
            lines.Add($"{s.Name},{s.FirstTime ?? "未打卡"},{s.Count},{string.Join("|", s.History)}");
        File.WriteAllLines(dlg.FileName, lines);
        StatusMessage = $"数据已导出到: {dlg.FileName}";
    }

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
    private void ClearAllRecords()
    {
        if (!VerifyAdminPwd("清空打卡记录")) return;
        for (int i = 3; i > 0; i--)
            if (!ModernDialog.Confirm($"确定要清空所有打卡记录？此操作不可恢复！\n还有 {i} 次警告", "警告"))
                return;
        if (!ModernDialog.Confirm("最后一次确认！确定要清空所有打卡记录？", "最终确认")) return;

        foreach (var stu in Students) { stu.FirstTime = null; stu.Count = 0; stu.History.Clear(); }
        PunchedCount = 0; UpdateRanking(); SaveAttendanceData();
        StatusMessage = "所有打卡记录已清空";
        if (IsOnline) _ = _server.SyncDataAsync(GetAttendanceDict());
    }

    private Dictionary<string, StudentAttendance> GetAttendanceDict() =>
        Students.ToDictionary(s => s.Name, s => new StudentAttendance
        { Name = s.Name, Count = s.Count, FirstTime = s.FirstTime, History = s.History });

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

// ---- 排名项 ----
public class RankingItem
{
    public int Rank { get; set; }
    public string Name { get; set; } = "";
    public string Time { get; set; } = "";
}

// ---- 对话框接口 ----
public interface IDialogVm
{
    Action? CloseAction { get; set; }
    Action? SaveAction { get; set; }
}

public class RemoteSettingsVm : IDialogVm
{
    public string Ip { get; set; } = "";
    public int Port { get; set; } = 5000;
    public string Password { get; set; } = "";
    public Action? CloseAction { get; set; }
    public Action? SaveAction { get; set; }
}

public class AdminSettingsVm : IDialogVm
{
    public string School { get; set; } = "";
    public string Nj { get; set; } = "";
    public string ClassId { get; set; } = "";
    public string Km { get; set; } = "";
    public int ButtonRows { get; set; } = 6;
    public int ButtonCols { get; set; } = 6;
    public string NewPassword { get; set; } = "";
    public string ConfirmPassword { get; set; } = "";
    public Action? CloseAction { get; set; }
    public Action? SaveAction { get; set; }
}
