using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Models;
using CheckIn.Client.Maui.Services;

namespace CheckIn.Client.Maui.ViewModels;

/// <summary>
/// 工作区配置（持久化到 workspace.json）
/// </summary>
public class WorkspaceConfig
{
    public List<TabInfo> Tabs { get; set; } = new();
    public string? ActiveTabId { get; set; }
}

/// <summary>
/// 任务树节点（用于左侧任务列表）
/// </summary>
public class TaskTreeNode : INotifyPropertyChanged
{
    public string Id { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public bool IsFolder { get; set; }
    public bool IsExpanded { get; set; } = true;
    public TaskTabViewModel? Tab { get; set; }
    public ObservableCollection<TaskTreeNode> Children { get; set; } = new();

    private bool _isSelected;
    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; OnPropertyChanged(); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

/// <summary>
/// 排名展示项数据模型
/// </summary>
public class RankingItem
{
    public int Rank { get; set; }
    public string Name { get; set; } = "";
    public string Time { get; set; } = "";
}

public class MainViewModel : INotifyPropertyChanged
{
    /// <summary>全局应用配置</summary>
    public AppConfig Config { get; } = new();

    private readonly string _baseDir;
    private readonly string _configFile;
    private readonly string _workspaceFile;

    /// <summary>当前打开的标签页集合</summary>
    public ObservableCollection<TaskTabViewModel> Tabs { get; } = new();

    /// <summary>任务树节点（左侧列表）</summary>
    public ObservableCollection<TaskTreeNode> TaskTree { get; } = new();

    /// <summary>后台任务列表（未打开的任务也在后台连接服务器）</summary>
    private List<TaskTabViewModel> _backgroundTasks = new();

    private TaskTabViewModel? _activeTab;
    /// <summary>当前活跃的标签页</summary>
    public TaskTabViewModel? ActiveTab
    {
        get => _activeTab;
        set
        {
            _activeTab = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(WindowTitle));
            OnPropertyChanged(nameof(HasActiveTab));
        }
    }

    /// <summary>是否有活跃标签页</summary>
    public bool HasActiveTab => ActiveTab != null;

    /// <summary>窗口标题（含版本号和作者信息）</summary>
    public string WindowTitle
    {
        get
        {
            var tabName = ActiveTab != null ? $" - {ActiveTab.TabDisplayName}" : "";
            return $"AgoraIn{tabName} {AppConfig.Version}  作者: 刘宇晨";
        }
    }

    // ---- 全局 Commands ----
    public ICommand AddTabCommand { get; }
    public ICommand CloseTabCommand { get; }
    public ICommand ShowRemoteSettingsCommand { get; }
    public ICommand CheckServerStatusCommand { get; }
    public ICommand LoadFromServerCommand { get; }
    public ICommand SyncToServerCommand { get; }
    public ICommand ShowAdminSettingsCommand { get; }
    public ICommand OpenGithubCommand { get; }
    public ICommand CheckVersionCommand { get; }
    public ICommand ShowAboutCommand { get; }
    public ICommand CreateSignInCommand { get; }
    public ICommand ShowSettingsCommand { get; }

    private readonly ServerService _serverService;

    public MainViewModel(ServerService serverService)
    {
        _serverService = serverService;
        _baseDir = FileSystem.AppDataDirectory;
        _configFile = Path.Combine(_baseDir, "config.json");
        _workspaceFile = Path.Combine(_baseDir, "workspace.json");

        LoadGlobalConfig();

        AddTabCommand = new RelayCommand(_ => AddTab());
        CloseTabCommand = new RelayCommand(CloseTab);
        ShowRemoteSettingsCommand = new RelayCommand(async _ => await ShowRemoteSettingsAsync());
        CheckServerStatusCommand = new RelayCommand(async _ => await CheckServerStatusAsync());
        LoadFromServerCommand = new RelayCommand(async _ => await LoadFromServerAsync());
        SyncToServerCommand = new RelayCommand(async _ => await SyncToServerAsync());
        ShowAdminSettingsCommand = new RelayCommand(async _ => await ShowAdminSettingsAsync());
        OpenGithubCommand = new RelayCommand(_ => OpenUrl("https://github.com/liuyuchen012/daikai"));
        CheckVersionCommand = new RelayCommand(_ => OpenUrl("https://github.com/liuyuchen012/daikai/releases"));
        ShowAboutCommand = new RelayCommand(async _ => await ShowAboutAsync());
        CreateSignInCommand = new RelayCommand(async _ => await CreateSignInAsync());
        ShowSettingsCommand = new RelayCommand(async _ => await Shell.Current.GoToAsync("Settings"));

        LoadWorkspace();
        InitializeBackgroundTasks();
    }

    // ---- 任务树管理 ----
    public void RefreshTaskTree()
    {
        TaskTree.Clear();

        var rootNode = new TaskTreeNode
        {
            Id = "root",
            DisplayName = "我的任务",
            IsFolder = true,
            IsExpanded = true
        };

        var allTasks = new List<TaskTabViewModel>();
        allTasks.AddRange(Tabs);
        allTasks.AddRange(_backgroundTasks);

        foreach (var tab in allTasks.OrderBy(t => t.Config.Name))
        {
            var node = new TaskTreeNode
            {
                Id = tab.TabId,
                DisplayName = tab.TabDisplayName,
                IsFolder = false,
                Tab = tab
            };
            rootNode.Children.Add(node);
        }

        TaskTree.Add(rootNode);
    }

    public void OnTaskTreeSelected(TaskTreeNode? node)
    {
        if (node == null || node.IsFolder) return;
        if (node.Tab == null) return;

        if (!Tabs.Contains(node.Tab))
        {
            Tabs.Add(node.Tab);
        }
        ActiveTab = node.Tab;
        SaveWorkspace();
    }

    // ---- 工作区管理 ----
    public void AddTab(string? name = null, string? km = null, bool isSignIn = false, string? signInTaskId = null)
    {
        var id = Guid.NewGuid().ToString("N")[..8];
        var tab = new TaskTabViewModel(id, _baseDir, Config, _serverService);
        if (!string.IsNullOrEmpty(name)) tab.Config.Name = name;
        if (!string.IsNullOrEmpty(km)) tab.Config.Km = km;
        if (isSignIn)
        {
            tab.Config.IsSignInTask = true;
            tab.Config.SignInTaskId = signInTaskId;
        }
        if (string.IsNullOrEmpty(tab.Config.Name)) tab.Config.Name = $"任务{Tabs.Count + 1}";
        tab.SaveConfig();

        if (isSignIn && !string.IsNullOrEmpty(signInTaskId))
        {
            tab.UpdateGlobalServerConfig(Config.ServerIp, Config.ServerPort, Config.ServerPassword);
        }

        Tabs.Add(tab);
        ActiveTab = tab;
        SaveWorkspace();
        RefreshTaskTree();
        OnPropertyChanged(nameof(WindowTitle));
    }

    private async void CloseTab(object? param)
    {
        if (param is not TaskTabViewModel tab) return;

        bool confirm = await DialogHelper.ConfirmAsync($"确定关闭「{tab.TabDisplayName}」？\n数据将保留在磁盘，可从左侧任务列表重新打开。");
        if (!confirm) return;

        int idx = Tabs.IndexOf(tab);
        Tabs.Remove(tab);

        if (!_backgroundTasks.Contains(tab))
        {
            _backgroundTasks.Add(tab);
        }

        if (Tabs.Count > 0)
        {
            ActiveTab = Tabs[Math.Min(idx, Tabs.Count - 1)];
        }
        else
        {
            ActiveTab = null;
        }
        SaveWorkspace();
        OnPropertyChanged(nameof(WindowTitle));
    }

    public void SwitchToTab(TaskTabViewModel tab)
    {
        if (!Tabs.Contains(tab))
        {
            Tabs.Add(tab);
            _backgroundTasks.Remove(tab);
        }
        ActiveTab = tab;
        SaveWorkspace();
    }

    public void RenameTab(TaskTabViewModel tab, string newName)
    {
        if (string.IsNullOrWhiteSpace(newName)) return;
        tab.Config.Name = newName.Trim();
        tab.SaveConfig();
        tab.NotifyPropertyChanged(nameof(TaskTabViewModel.TabDisplayName));
        SaveWorkspace();
        RefreshTaskTree();
        OnPropertyChanged(nameof(WindowTitle));
    }

    public void DeleteTask(TaskTabViewModel tab)
    {
        if (Tabs.Contains(tab))
        {
            Tabs.Remove(tab);
        }

        _backgroundTasks.Remove(tab);

        tab.Dispose();

        var tabDir = Path.Combine(_baseDir, "data", "tabs", tab.TabId);
        if (Directory.Exists(tabDir))
        {
            Directory.Delete(tabDir, recursive: true);
        }

        if (ActiveTab == tab)
        {
            ActiveTab = Tabs.FirstOrDefault() ?? _backgroundTasks.FirstOrDefault();
        }

        SaveWorkspace();
        RefreshTaskTree();
        OnPropertyChanged(nameof(WindowTitle));
    }

    private void InitializeBackgroundTasks()
    {
        var tabsDir = Path.Combine(_baseDir, "data", "tabs");
        if (!Directory.Exists(tabsDir)) return;

        foreach (var tabDir in Directory.GetDirectories(tabsDir))
        {
            var tabId = Path.GetFileName(tabDir);
            if (Tabs.Any(t => t.TabId == tabId)) continue;

            var bgTab = TaskTabViewModel.CreateBackgroundInstance(tabId, _baseDir, Config, _serverService);
            _backgroundTasks.Add(bgTab);
        }
    }

    private void LoadWorkspace()
    {
        if (File.Exists(_workspaceFile))
        {
            try
            {
                var ws = JsonSerializer.Deserialize<WorkspaceConfig>(File.ReadAllText(_workspaceFile));
                if (ws != null && ws.Tabs.Count > 0)
                {
                    foreach (var info in ws.Tabs)
                    {
                        var tab = new TaskTabViewModel(info.Id, _baseDir, Config, _serverService);
                        Tabs.Add(tab);
                    }
                    ActiveTab = Tabs.FirstOrDefault(t => t.TabId == ws.ActiveTabId) ?? Tabs.FirstOrDefault();
                    RefreshTaskTree();
                    return;
                }
            }
            catch { }
        }

        // 全新安装：创建默认标签
        AddTab("默认任务", "数学");
    }

    public void SaveWorkspace()
    {
        var ws = new WorkspaceConfig
        {
            Tabs = Tabs.Select(t => new TabInfo { Id = t.TabId, Name = t.TabDisplayName }).ToList(),
            ActiveTabId = ActiveTab?.TabId
        };
        File.WriteAllText(_workspaceFile, JsonSerializer.Serialize(ws, new JsonSerializerOptions { WriteIndented = true }));
    }

    // ---- 全局配置 ----
    private void LoadGlobalConfig()
    {
        if (!File.Exists(_configFile)) return;
        try
        {
            var c = JsonSerializer.Deserialize<AppConfig>(File.ReadAllText(_configFile));
            if (c != null)
            {
                Config.OnlineMode = c.OnlineMode;
                Config.ServerIp = c.ServerIp;
                Config.ServerPort = c.ServerPort;
                Config.ServerPassword = c.ServerPassword;
                Config.AdminPasswordHash = c.AdminPasswordHash;
            }
        }
        catch { }
    }

    public void SaveGlobalConfig()
    {
        File.WriteAllText(_configFile, JsonSerializer.Serialize(Config, new JsonSerializerOptions { WriteIndented = true }));
    }

    // ---- 远程设置 ----
    private async Task ShowRemoteSettingsAsync()
    {
        string? ip = await DialogHelper.PromptAsync("服务器地址", "请输入服务器IP地址:", Config.ServerIp);
        if (ip == null) return;

        string? portStr = await DialogHelper.PromptAsync("服务器端口", "请输入服务器端口:", Config.ServerPort.ToString());
        if (portStr == null) return;
        if (!int.TryParse(portStr, out int port)) return;

        string? password = await DialogHelper.PromptAsync("服务器密码", "请输入服务器密码:", Config.ServerPassword);
        if (password == null) return;

        Config.ServerIp = ip;
        Config.ServerPort = port;
        Config.ServerPassword = password;
        SaveGlobalConfig();

        foreach (var tab in Tabs)
            tab.UpdateGlobalServerConfig(ip, port, password);
        foreach (var bgTab in _backgroundTasks)
            bgTab.UpdateGlobalServerConfig(ip, port, password);

        StatusMessage = "远程设置已保存";
    }

    private string _statusMessage = "就绪";
    public string StatusMessage
    {
        get => _statusMessage;
        set
        {
            _statusMessage = value;
            OnPropertyChanged();
            if (ActiveTab != null) ActiveTab.StatusMessage = value;
        }
    }

    // ---- 服务器操作 ----
    private async Task CheckServerStatusAsync()
    {
        if (ActiveTab == null) return;
        await ActiveTab.CheckServerStatusAsync();
    }

    private async Task LoadFromServerAsync()
    {
        if (ActiveTab == null) return;
        await ActiveTab.LoadFromServerAsync();
    }

    private async Task SyncToServerAsync()
    {
        if (ActiveTab == null) return;
        await ActiveTab.SyncToServerAsync();
    }

    // ---- 管理员设置 ----
    private async Task ShowAdminSettingsAsync()
    {
        if (!await VerifyAdminPwd("访问管理员设置")) return;

        string? newPwd = await DialogHelper.PromptAsync("管理员设置", "新密码(留空不改):", "", true);
        if (newPwd == null) return;

        if (!string.IsNullOrEmpty(newPwd))
        {
            string? confirmPwd = await DialogHelper.PromptAsync("管理员设置", "确认密码:", "", true);
            if (confirmPwd == null) return;
            if (newPwd != confirmPwd)
            {
                await DialogHelper.AlertAsync("两次输入的密码不一致", "错误");
                return;
            }
            Config.AdminPasswordHash = HashPassword(newPwd);
        }

        SaveGlobalConfig();
        await DialogHelper.AlertAsync("全局设置已保存并生效");
    }

    // ---- 创建签到 ----
    private async Task CreateSignInAsync()
    {
        var page = new Pages.CreateSignInPage(_serverService, this);
        await Shell.Current.Navigation.PushModalAsync(new NavigationPage(page));
    }

    // ---- 关于 ----
    private async Task ShowAboutAsync()
    {
        await DialogHelper.AlertAsync(
            $"AgoraIn {AppConfig.Version}\n\n" +
            "开发者: 刘宇晨\n" +
            "联系邮箱: liuyuchen032901@outlook.com\n\n" +
            "(c) 2026 保留全部权利",
            "关于");
    }

    // ---- 辅助方法 ----
    private static void OpenUrl(string url)
    {
        if (OperatingSystem.IsIOS()) return;
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
        catch { }
    }

    private async Task<bool> VerifyAdminPwd(string action)
    {
        if (string.IsNullOrEmpty(Config.AdminPasswordHash)) return true;
        var pwd = await DialogHelper.PromptAsync(action, "请输入管理员密码:", "", true);
        return pwd != null && HashPassword(pwd) == Config.AdminPasswordHash;
    }

    private static string HashPassword(string pwd)
    {
        using var sha = System.Security.Cryptography.SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pwd)));
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
