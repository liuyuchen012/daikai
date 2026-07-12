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

namespace CheckIn.Client.ViewModels;

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

public class MainViewModel : INotifyPropertyChanged
{
    /// <summary>全局应用配置</summary>
    public AppConfig Config { get; } = new();

    private readonly string _baseDir;
    private readonly string _configFile;
    private readonly string _workspaceFile;

    /// <summary>当前打开的标签页集合</summary>
    public ObservableCollection<TaskTabViewModel> Tabs { get; } = new();

    /// <summary>
    /// 任务树节点（左侧列表）
    /// </summary>
    public ObservableCollection<TaskTreeNode> TaskTree { get; } = new();

    /// <summary>
    /// 后台任务列表（未打开的任务也在后台连接服务器）
    /// </summary>
    private List<TaskTabViewModel> _backgroundTasks = new();

    private TaskTabViewModel? _activeTab;
    /// <summary>当前活跃的标签页</summary>
    public TaskTabViewModel? ActiveTab
    {
        get => _activeTab;
        set { _activeTab = value; OnPropertyChanged(); OnPropertyChanged(nameof(WindowTitle)); OnPropertyChanged(nameof(HasActiveTab)); }
    }

    /// <summary>是否有活跃标签页（用于控制占位提示的显示）</summary>
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
    /// <summary>添加新标签页命令</summary>
    public ICommand AddTabCommand { get; }
    /// <summary>关闭标签页命令</summary>
    public ICommand CloseTabCommand { get; }
    /// <summary>退出应用命令</summary>
    public ICommand ExitCommand { get; }
    /// <summary>显示远程设置对话框命令</summary>
    public ICommand ShowRemoteSettingsCommand { get; }
    /// <summary>检查服务器状态命令</summary>
    public ICommand CheckServerStatusCommand { get; }
    /// <summary>从服务器加载数据命令</summary>
    public ICommand LoadFromServerCommand { get; }
    /// <summary>同步数据到服务器命令</summary>
    public ICommand SyncToServerCommand { get; }
    /// <summary>显示管理员设置对话框命令</summary>
    public ICommand ShowAdminSettingsCommand { get; }
    /// <summary>打开 GitHub 仓库命令</summary>
    public ICommand OpenGithubCommand { get; }
    /// <summary>检查版本列表命令</summary>
    public ICommand CheckVersionCommand { get; }
    /// <summary>显示关于对话框命令</summary>
    public ICommand ShowAboutCommand { get; }

    /// <summary>
    /// 构造函数：初始化应用目录、加载配置、恢复工作区、初始化后台任务
    /// </summary>
    public MainViewModel()
    {
        _baseDir = AppDomain.CurrentDomain.BaseDirectory;
        _configFile = Path.Combine(_baseDir, "config.json");
        _workspaceFile = Path.Combine(_baseDir, "workspace.json");

        LoadGlobalConfig();

        AddTabCommand = new RelayCommand(_ => AddTab());
        CloseTabCommand = new RelayCommand(CloseTab);
        ExitCommand = new RelayCommand(_ => Application.Current.Shutdown());
        ShowRemoteSettingsCommand = new RelayCommand(_ => ShowRemoteSettings());
        CheckServerStatusCommand = new RelayCommand(async _ => await CheckServerStatusAsync());
        LoadFromServerCommand = new RelayCommand(async _ => await LoadFromServerAsync());
        SyncToServerCommand = new RelayCommand(async _ => await SyncToServerAsync());
        ShowAdminSettingsCommand = new RelayCommand(_ => ShowAdminSettings());
        OpenGithubCommand = new RelayCommand(_ => OpenUrl("https://github.com/liuyuchen012/daikai"));
        CheckVersionCommand = new RelayCommand(_ => OpenUrl("https://github.com/liuyuchen012/daikai/releases"));
        ShowAboutCommand = new RelayCommand(_ => ShowAbout());

        LoadWorkspace();
        InitializeBackgroundTasks();
    }

    // ---- 任务树管理 ----
    /// <summary>
    /// 重建任务树：根节点"我的任务"下包含所有已打开和后台的任务
    /// </summary>
    public void RefreshTaskTree()
    {
        TaskTree.Clear();

        // 根节点：本机任务
        var rootNode = new TaskTreeNode
        {
            Id = "root",
            DisplayName = "我的任务",
            IsFolder = true,
            IsExpanded = true
        };

        // 收集所有任务（已打开的 + 后台的）
        var allTasks = new List<TaskTabViewModel>();
        allTasks.AddRange(Tabs);
        allTasks.AddRange(_backgroundTasks);

        // 按名称排序
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

    /// <summary>
    /// 任务树节点选中：如果标签未打开则先打开，然后设置为活跃标签
    /// </summary>
    public void OnTaskTreeSelected(TaskTreeNode? node)
    {
        if (node == null || node.IsFolder) return;
        if (node.Tab == null) return;

        // 打开标签页
        if (!Tabs.Contains(node.Tab))
        {
            Tabs.Add(node.Tab);
        }
        ActiveTab = node.Tab;
        SaveWorkspace();
    }

    // ---- 工作区管理 ----
    /// <summary>
    /// 添加新标签页：生成 8 位短 ID、创建 ViewModel、保存配置并刷新任务树
    /// </summary>
    public void AddTab(string? name = null, string? km = null, bool isSignIn = false, string? signInTaskId = null)
    {
        var id = Guid.NewGuid().ToString("N")[..8];
        var tab = new TaskTabViewModel(id, _baseDir, Config);
        if (!string.IsNullOrEmpty(name)) tab.Config.Name = name;
        if (!string.IsNullOrEmpty(km)) tab.Config.Km = km;
        if (isSignIn)
        {
            tab.Config.IsSignInTask = true;
            tab.Config.SignInTaskId = signInTaskId;
        }
        if (string.IsNullOrEmpty(tab.Config.Name)) tab.Config.Name = $"任务{Tabs.Count + 1}";
        tab.SaveConfig();
        // 如果是签到任务，重新初始化服务器连接以使用正确的 TaskId
        if (isSignIn && !string.IsNullOrEmpty(signInTaskId))
        {
            tab.UpdateGlobalServerConfig(Config.ServerIp, Config.ServerPort, Config.ServerPassword, AppConfig.Version);
        }
        Tabs.Add(tab);
        ActiveTab = tab;
        SaveWorkspace();
        RefreshTaskTree();
        OnPropertyChanged(nameof(WindowTitle));
    }

    /// <summary>
    /// 关闭标签页：从 Tabs 移除并加入后台任务列表，保持服务器连接
    /// </summary>
    private void CloseTab(object? param)
    {
        if (param is not TaskTabViewModel tab) return;
        if (!ModernDialog.Confirm($"确定关闭「{tab.TabDisplayName}」？\n数据将保留在磁盘，可从左侧任务列表重新打开。")) return;

        // 不要 Dispose，保持后台连接
        int idx = Tabs.IndexOf(tab);
        Tabs.Remove(tab);

        // 添加到后台任务列表
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
        // 注意：不调用 RefreshTaskTree()，任务树保持不变
        OnPropertyChanged(nameof(WindowTitle));
    }

    /// <summary>
    /// 切换到指定标签页（如果已关闭则重新打开并移出后台列表）
    /// </summary>
    public void SwitchToTab(TaskTabViewModel tab)
    {
        // 如果标签页已关闭，重新添加到 Tabs
        if (!Tabs.Contains(tab))
        {
            Tabs.Add(tab);
            // 从后台列表移除
            _backgroundTasks.Remove(tab);
        }
        ActiveTab = tab;
        SaveWorkspace();
    }

    /// <summary>
    /// 重命名标签页：更新配置中的名称，刷新任务树和窗口标题
    /// </summary>
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

    /// <summary>
    /// 删除任务（包括所有数据）
    /// </summary>
    public void DeleteTask(TaskTabViewModel tab)
    {
        // 从 Tabs 移除
        if (Tabs.Contains(tab))
        {
            Tabs.Remove(tab);
        }

        // 从后台任务列表移除
        _backgroundTasks.Remove(tab);

        // Dispose 释放资源
        tab.Dispose();

        // 删除任务目录
        var tabDir = Path.Combine(_baseDir, "data", "tabs", tab.TabId);
        if (Directory.Exists(tabDir))
        {
            Directory.Delete(tabDir, recursive: true);
        }

        // 更新活跃标签页
        if (ActiveTab == tab)
        {
            ActiveTab = Tabs.FirstOrDefault() ?? _backgroundTasks.FirstOrDefault();
        }

        SaveWorkspace();
        RefreshTaskTree();
        OnPropertyChanged(nameof(WindowTitle));
    }

    /// <summary>
    /// 初始化后台任务（未打开的任务也在后台连接服务器）
    /// </summary>
    private void InitializeBackgroundTasks()
    {
        // 扫描 data/tabs 目录下所有任务
        var tabsDir = Path.Combine(_baseDir, "data", "tabs");
        if (!Directory.Exists(tabsDir)) return;

        foreach (var tabDir in Directory.GetDirectories(tabsDir))
        {
            var tabId = Path.GetFileName(tabDir);
            // 跳过已打开的标签页
            if (Tabs.Any(t => t.TabId == tabId)) continue;

            // 创建后台任务实例（轻量级，不加载UI数据）
            var bgTab = TaskTabViewModel.CreateBackgroundInstance(tabId, _baseDir, Config);
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
                        var tab = new TaskTabViewModel(info.Id, _baseDir, Config);
                        Tabs.Add(tab);
                    }
                    ActiveTab = Tabs.FirstOrDefault(t => t.TabId == ws.ActiveTabId) ?? Tabs.FirstOrDefault();
                    RefreshTaskTree();
                    return;
                }
            }
            catch { }
        }

        // 向后兼容：检查旧版 config.json
        if (File.Exists(_configFile))
        {
            try
            {
                var oldConfig = JsonSerializer.Deserialize<AppConfig>(File.ReadAllText(_configFile));
                if (oldConfig != null && !string.IsNullOrEmpty(oldConfig.School))
                {
                    var tabName = $"{oldConfig.School}{oldConfig.Nj}年{oldConfig.ClassId}班";
                    // 迁移旧数据到新标签
                    var id = "legacy01";
                    var tabDir = Path.Combine(_baseDir, "data", "tabs", id);
                    Directory.CreateDirectory(tabDir);

                    // 迁移数据文件
                    var oldDataFile = Path.Combine(_baseDir, "attendance.dat");
                    var oldNameFile = Path.Combine(_baseDir, "name.txt");
                    if (File.Exists(oldDataFile) && !File.Exists(Path.Combine(tabDir, "attendance.dat")))
                        File.Copy(oldDataFile, Path.Combine(tabDir, "attendance.dat"), true);
                    if (File.Exists(oldNameFile) && !File.Exists(Path.Combine(tabDir, "name.txt")))
                        File.Copy(oldNameFile, Path.Combine(tabDir, "name.txt"), true);

                    // 写入标签配置
                    var tabConfig = new TabConfig
                    {
                        Name = tabName,
                        Km = oldConfig.Km,
                        ButtonRows = oldConfig.ButtonRows,
                        ButtonCols = oldConfig.ButtonCols,
                        OnlineMode = oldConfig.OnlineMode
                    };
                    File.WriteAllText(Path.Combine(tabDir, "config.json"),
                        JsonSerializer.Serialize(tabConfig, new JsonSerializerOptions { WriteIndented = true }));

                    var tab = new TaskTabViewModel(id, _baseDir, Config);
                    Tabs.Add(tab);
                    ActiveTab = tab;
                    SaveWorkspace();
                    RefreshTaskTree();
                    return;
                }
            }
            catch { }
        }

        // 全新安装：创建一个默认标签
        AddTab("默认任务", "数学");
    }

    private void SaveWorkspace()
    {
        var ws = new WorkspaceConfig
        {
            Tabs = Tabs.Select(t => new TabInfo { Id = t.TabId, Name = t.TabDisplayName }).ToList(),
            ActiveTabId = ActiveTab?.TabId
        };
        File.WriteAllText(_workspaceFile, JsonSerializer.Serialize(ws, new JsonSerializerOptions { WriteIndented = true }));
    }

    // ---- 全局配置 ----
    /// <summary>
    /// 从 config.json 加载全局配置（在线模式、服务器地址等），跳过已迁移到标签页的字段
    /// </summary>
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
                // 版本号不从文件加载，始终使用编译时定义的版本
            }
        }
        catch { }
    }

    /// <summary>
    /// 保存全局配置到 config.json
    /// </summary>
    public void SaveGlobalConfig()
    {
        File.WriteAllText(_configFile, JsonSerializer.Serialize(Config, new JsonSerializerOptions { WriteIndented = true }));
    }

    // ---- 远程设置 ----
    /// <summary>
    /// 显示远程服务器设置对话框，保存后同步到所有标签页和后台任务
    /// </summary>
    private void ShowRemoteSettings()
    {
        var vm = new RemoteSettingsVm
        {
            Ip = Config.ServerIp, Port = Config.ServerPort, Password = Config.ServerPassword
        };
        vm.SaveAction = () =>
        {
            Config.ServerIp = vm.Ip; Config.ServerPort = vm.Port; Config.ServerPassword = vm.Password;
            SaveGlobalConfig();
            // 同步到所有标签页和后台任务
            foreach (var tab in Tabs)
                tab.UpdateGlobalServerConfig(vm.Ip, vm.Port, vm.Password, AppConfig.Version);
            foreach (var bgTab in _backgroundTasks)
                bgTab.UpdateGlobalServerConfig(vm.Ip, vm.Port, vm.Password, AppConfig.Version);
            StatusMessage = "远程设置已保存";
        };
        ShowDialog("远程服务器设置", vm, CreateRemoteSettingsView);
    }

    private string _statusMessage = "就绪";
    /// <summary>全局状态栏消息（会同步到活跃标签页）</summary>
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

    // ---- 服务器操作（委托给活跃标签） ----
    /// <summary>检查服务器在线状态（委托给当前活跃标签页）</summary>
    private async Task CheckServerStatusAsync()
    {
        if (ActiveTab == null) return;
        await ActiveTab.CheckServerStatusAsync();
    }

    /// <summary>从服务器加载数据（委托给当前活跃标签页）</summary>
    private async Task LoadFromServerAsync()
    {
        if (ActiveTab == null) return;
        await ActiveTab.LoadFromServerAsync();
    }

    /// <summary>同步数据到服务器（委托给当前活跃标签页）</summary>
    private async Task SyncToServerAsync()
    {
        if (ActiveTab == null) return;
        await ActiveTab.SyncToServerAsync();
    }

    // ---- 全局管理员设置 ----
    /// <summary>
    /// 显示全局管理员设置对话框（修改管理员密码），需要密码验证
    /// </summary>
    private void ShowAdminSettings()
    {
        if (!VerifyAdminPwd("访问管理员设置")) return;
        var vm = new AdminSettingsVm
        {
            NewPassword = "", ConfirmPassword = ""
        };
        vm.SaveAction = () =>
        {
            if (!string.IsNullOrEmpty(vm.NewPassword))
            {
                if (vm.NewPassword != vm.ConfirmPassword)
                {
                    ModernDialog.Alert("两次输入的密码不一致", "错误");
                    return;
                }
                Config.AdminPasswordHash = HashPassword(vm.NewPassword);
            }
            SaveGlobalConfig();
            ModernDialog.Alert("全局设置已保存并生效");
        };
        ShowDialog("管理员设置", vm, CreateAdminSettingsView);
    }

    // ---- 关于 ----
    private void ShowAbout()
    {
        var logoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo.svg");
        var aboutWin = new Window
        {
            Title = "关于", Width = 360, Height = 320,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false,
            Content = new System.Windows.Controls.StackPanel
            {
                Margin = new Thickness(10),
                Children =
                {
                    new SharpVectors.Converters.SvgViewbox
                    {
                        Source = new Uri(logoPath),
                        Width = 80, Height = 80,
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                        Margin = new Thickness(0, 10, 0, 10)
                    },
                    new System.Windows.Controls.TextBlock { Text = $"版本: {AppConfig.Version}", FontSize = 14, Margin = new Thickness(0,0,0,5) },
                    new System.Windows.Controls.TextBlock { Text = "开发者: 刘宇晨", FontSize = 14, Margin = new Thickness(0,0,0,5) },
                    new System.Windows.Controls.TextBlock { Text = "联系邮箱: liuyuchen032901@outlook.com", FontSize = 14, Margin = new Thickness(0,0,0,5) },
                    new System.Windows.Controls.TextBlock { Text = "© 2026 保留全部权利", FontSize = 12, Foreground = System.Windows.Media.Brushes.Gray, Margin = new Thickness(0,20,0,0) }
                }
            }
        };
        aboutWin.ShowDialog();
    }

    // ---- 辅助方法 ----
    /// <summary>使用系统默认浏览器打开 URL</summary>
    private static void OpenUrl(string url)
    {
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
        catch { }
    }

    /// <summary>验证管理员密码（如果未设置密码则直接通过）</summary>
    private bool VerifyAdminPwd(string action)
    {
        if (string.IsNullOrEmpty(Config.AdminPasswordHash)) return true;
        var pwd = AskPassword($"需要管理员权限执行: {action}");
        return pwd != null && HashPassword(pwd) == Config.AdminPasswordHash;
    }

    /// <summary>对密码进行 SHA256 哈希</summary>
    private static string HashPassword(string pwd)
    {
        using var sha = System.Security.Cryptography.SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pwd)));
    }

    /// <summary>弹出密码输入对话框，返回输入的密码</summary>
    private string? AskPassword(string title)
    {
        string? result = null;
        var win = new Window
        {
            Title = title, Width = 320, Height = 180,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false
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

    /// <summary>
    /// 通用的对话框显示方法，绑定 ViewModel 并支持 CloseAction
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
    /// 构建远程设置对话框 UI 视图
    /// </summary>
    private FrameworkElement CreateRemoteSettingsView(RemoteSettingsVm? vm)
    {
        if (vm == null) return new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };
        var panel = new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };
        AddFieldRow(panel, "服务器地址:", vm.Ip, v => vm.Ip = v);
        AddFieldRow(panel, "服务器端口:", vm.Port.ToString(), v => { if (int.TryParse(v, out var n)) vm.Port = n; });
        AddFieldRow(panel, "服务器密码:", vm.Password, v => vm.Password = v, true);
        AddSaveCancel(panel, vm);
        return panel;
    }

    /// <summary>
    /// 构建管理员设置对话框 UI 视图
    /// </summary>
    private FrameworkElement CreateAdminSettingsView(AdminSettingsVm? vm)
    {
        if (vm == null) return new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };
        var panel = new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };
        AddFieldRow(panel, "新密码(留空不改):", vm.NewPassword, v => vm.NewPassword = v, true);
        AddFieldRow(panel, "确认密码:", vm.ConfirmPassword, v => vm.ConfirmPassword = v, true);
        AddSaveCancel(panel, vm);
        return panel;
    }

    /// <summary>
    /// 向面板添加一行表单字段（标签 + 文本框/密码框）
    /// </summary>
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

    /// <summary>
    /// 向面板添加保存/取消按钮行
    /// </summary>
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

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string? name = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

/// <summary>
/// 排名展示项数据模型
/// </summary>
public class RankingItem
{
    /// <summary>排名序号</summary>
    public int Rank { get; set; }
    /// <summary>学生姓名</summary>
    public string Name { get; set; } = "";
    /// <summary>打卡时间（HH:mm 格式）</summary>
    public string Time { get; set; } = "";
}

/// <summary>
/// 对话框 ViewModel 接口，提供关闭和保存的回调委托
/// </summary>
public interface IDialogVm
{
    /// <summary>关闭对话框的回调</summary>
    Action? CloseAction { get; set; }
    /// <summary>保存设置的回调</summary>
    Action? SaveAction { get; set; }
}

/// <summary>
/// 远程服务器设置对话框的 ViewModel
/// </summary>
public class RemoteSettingsVm : IDialogVm
{
    /// <summary>服务器 IP 地址</summary>
    public string Ip { get; set; } = "";
    /// <summary>服务器端口</summary>
    public int Port { get; set; } = 5250;
    /// <summary>服务器连接密码</summary>
    public string Password { get; set; } = "";
    public Action? CloseAction { get; set; }
    public Action? SaveAction { get; set; }
}

/// <summary>
/// 全局管理员设置对话框的 ViewModel
/// </summary>
public class AdminSettingsVm : IDialogVm
{
    /// <summary>新管理员密码（留空表示不修改）</summary>
    public string NewPassword { get; set; } = "";
    /// <summary>确认密码（需与新密码一致）</summary>
    public string ConfirmPassword { get; set; } = "";
    public Action? CloseAction { get; set; }
    public Action? SaveAction { get; set; }
}
