using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Net.Http.Json;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Input;
using CheckIn.Client.Models;
using CheckIn.Client.Services;
using CheckIn.Shared.Models;

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
    /// <summary>当前活跃的标签页（L7：切换时同步更新各标签的 IsActive 用于标签栏高亮）</summary>
    public TaskTabViewModel? ActiveTab
    {
        get => _activeTab;
        set
        {
            if (_activeTab != value)
            {
                if (_activeTab != null) _activeTab.IsActive = false;
                _activeTab = value;
                if (_activeTab != null) _activeTab.IsActive = true;
            }
            OnPropertyChanged(); OnPropertyChanged(nameof(WindowTitle)); OnPropertyChanged(nameof(HasActiveTab));
        }
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
    /// <summary>显示学生列表管理对话框命令（大屏模式设置菜单）</summary>
    public ICommand ShowStudentListCommand { get; }
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
        ShowStudentListCommand = new RelayCommand(_ => ShowStudentList());
        OpenGithubCommand = new RelayCommand(_ => OpenUrl("https://github.com/liuyuchen012/daikai"));
        CheckVersionCommand = new RelayCommand(async _ => await CheckForUpdateAsync());
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

        // 订阅服务端推送任务事件
        tab.PendingTasksReceived += HandlePendingTasks;
        // 订阅集控平台新版本通知（管理员弹窗）
        tab.ServerUpdateAvailable += HandleServerUpdate;
        // 订阅集控平台呼叫通知（待下课 / 应急 / 传唤）
        tab.CallReceived += HandleCallReceived;

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
    /// 处理服务端推送的新任务配置：创建对应的签到任务标签页
    /// </summary>
    private void HandlePendingTasks(List<PendingTaskConfig> pendingTasks)
    {
        foreach (var pt in pendingTasks)
        {
            // 检查是否已存在相同 TaskId 的任务
            var existingTab = Tabs.FirstOrDefault(t => t.Config.SignInTaskId == pt.TaskId);
            if (existingTab == null)
            {
                existingTab = _backgroundTasks.FirstOrDefault(t => t.Config.SignInTaskId == pt.TaskId);
            }
            if (existingTab != null) continue; // 已存在，跳过

            // 创建新的签到任务标签页
            var taskName = string.IsNullOrEmpty(pt.TaskName) ? pt.Subject : pt.TaskName;
            var id = Guid.NewGuid().ToString("N")[..8];
            var tab = new TaskTabViewModel(id, _baseDir, Config);
            tab.Config.Name = taskName;
            tab.Config.Km = pt.Subject;
            tab.Config.IsSignInTask = true;
            tab.Config.SignInTaskId = pt.TaskId;
            tab.SaveConfig();

            // 导入学生名单
            if (pt.Students.Count > 0)
            {
                var namesPath = Path.Combine(_baseDir, "data", "tabs", id, "name.txt");
                File.WriteAllLines(namesPath, pt.Students);
                tab.LoadStudentNames();
            }

            // 订阅推送事件
            tab.PendingTasksReceived += HandlePendingTasks;
            // 订阅集控平台新版本通知（管理员弹窗）
            tab.ServerUpdateAvailable += HandleServerUpdate;
            // 订阅集控平台呼叫通知（待下课 / 应急 / 传唤）
            tab.CallReceived += HandleCallReceived;

            // 初始化服务器连接（使用后台模式）
            if (!string.IsNullOrEmpty(Config.ServerIp))
            {
                tab.UpdateGlobalServerConfig(Config.ServerIp, Config.ServerPort, Config.ServerPassword, AppConfig.Version);
            }

            // 添加到后台任务列表（不打开标签页）
            _backgroundTasks.Add(tab);
        }
    }

    /// <summary>
    /// 处理集控平台自身的新版本通知：向管理员用户弹窗提示
    /// </summary>
    private void HandleServerUpdate(string latestVersion, string downloadUrl)
    {
        ModernDialog.ServerUpdateAvailable(latestVersion, downloadUrl);
    }

    /// <summary>
    /// 处理集控平台下发的呼叫：大屏端醒目弹窗显示（待下课 / 应急 / 传唤）
    /// </summary>
    private void HandleCallReceived(CheckIn.Client.Models.CallMessage call)
    {
        ModernDialog.ShowCall(call);
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

                    // H3：若 legacy01 已完成迁移（配置已存在），直接复用，避免重复迁移/覆盖
                    if (File.Exists(Path.Combine(tabDir, "config.json")))
                    {
                        var existingTab = new TaskTabViewModel(id, _baseDir, Config);
                        Tabs.Add(existingTab);
                        ActiveTab = existingTab;
                        SaveWorkspace();
                        RefreshTaskTree();
                        return;
                    }

                    Directory.CreateDirectory(tabDir);

                    // 迁移数据文件（H3：失败时明确提示用户，不再静默吞掉异常）
                    try
                    {
                        var oldDataFile = Path.Combine(_baseDir, "attendance.dat");
                        var oldNameFile = Path.Combine(_baseDir, "name.txt");
                        if (File.Exists(oldDataFile))
                            File.Copy(oldDataFile, Path.Combine(tabDir, "attendance.dat"), false);
                        if (File.Exists(oldNameFile))
                            File.Copy(oldNameFile, Path.Combine(tabDir, "name.txt"), false);
                    }
                    catch (Exception ex)
                    {
                        ModernDialog.Alert($"旧版数据迁移失败：{ex.Message}\n\n旧数据仍保留在原位置（attendance.dat / name.txt），不会丢失，可手动处理。", "迁移提示");
                    }

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
    /// <summary>全局状态栏消息（M3：与活跃标签页的消息相互独立，不再强制覆盖）</summary>
    public string StatusMessage
    {
        get => _statusMessage;
        set { _statusMessage = value; OnPropertyChanged(); }
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

    // ---- 学生列表管理（大屏模式设置菜单）----
    /// <summary>
    /// 显示当前活跃标签页的学生列表管理对话框：支持添加、编辑（改姓名 + 勾选需要打卡的任务）、删除、批量删除
    /// </summary>
    private void ShowStudentList()
    {
        if (ActiveTab == null)
        {
            ModernDialog.Alert("请先打开或新建一个任务，再管理学生列表");
            return;
        }
        var tab = ActiveTab;
        var vm = new StudentListVm(tab, GetAllTasks());
        vm.AddAction = () =>
        {
            var name = AskText("添加学生", "请输入学生姓名:");
            if (string.IsNullOrWhiteSpace(name)) return;
            if (!tab.AddStudent(name.Trim()))
            {
                ModernDialog.Alert("添加失败：姓名不能为空或该学生已存在");
                return;
            }
            ModernDialog.Alert($"已添加学生: {name.Trim()}");
        };
        vm.EditAction = () => EditStudent(tab, vm.SelectedStudent);
        vm.DeleteAction = () =>
        {
            if (vm.SelectedStudent == null) { ModernDialog.Alert("请先选中要删除的学生"); return; }
            var stu = vm.SelectedStudent;
            if (!ModernDialog.Confirm($"确定删除学生「{stu.Name}」？\n该学生的打卡记录也会一并清除。", "删除确认")) return;
            if (tab.RemoveStudent(stu)) ModernDialog.Alert($"已删除学生: {stu.Name}");
        };
        vm.BatchDeleteAction = () =>
        {
            var selected = vm.SelectedStudents.Where(s => s != null).Distinct().ToList();
            if (selected.Count == 0) { ModernDialog.Alert("请先选中要删除的学生（可按住 Ctrl 或 Shift 多选）"); return; }
            var names = string.Join("、", selected.Select(s => s.Name));
            if (!ModernDialog.Confirm($"确定删除选中的 {selected.Count} 名学生？\n\n{names}\n\n他们的打卡记录也会一并清除。", "批量删除确认")) return;
            var removed = tab.RemoveStudents(selected);
            if (removed > 0)
            {
                ModernDialog.Alert($"已删除 {removed} 名学生");
                vm.SelectedStudents.Clear();
                vm.SelectedStudent = null;
            }
        };
        ShowDialog("学生列表", vm, CreateStudentListView, 520, 480);
    }

    /// <summary>
    /// 打开编辑学生对话框：可修改学生姓名，并勾选该学生需要进行打卡的任务
    /// </summary>
    private void EditStudent(TaskTabViewModel tab, StudentModel? stu)
    {
        if (stu == null) { ModernDialog.Alert("请先选中要编辑的学生"); return; }
        var dlg = new StudentListVm(tab, GetAllTasks())
        {
            IsEditMode = true,
            SelectedStudent = stu,
            NewName = stu.Name
        };
        // 初始化"需要进行打卡的任务"勾选状态：学生存在于某任务名单中即勾选；当前任务恒为勾选
        foreach (var t in dlg.AllTasks)
        {
            var contains = t == tab || ContainsStudent(t, stu.Name);
            dlg.TaskChecks[t.TabId] = contains;
        }
        dlg.SaveAction = () =>
        {
            var newName = dlg.NewName.Trim();
            if (string.IsNullOrEmpty(newName)) { ModernDialog.Alert("姓名不能为空"); return; }

            var oldName = stu.Name;
            // 1. 修改姓名（当前任务内，含打卡数据同步）
            if (oldName != newName && !tab.RenameStudent(stu, newName))
            {
                ModernDialog.Alert("改名失败：该姓名已存在或为空");
                return;
            }
            // 2. 同步到其他任务：
            //    - 先移除旧姓名（改名后旧姓名不再有效）
            //    - 勾选的任务加入新姓名，未勾选的任务不加
            foreach (var t in dlg.AllTasks)
            {
                if (t == tab) continue; // 当前任务已处理
                if (oldName != newName) SyncStudentToTask(t, oldName, false);
                var checked_ = dlg.TaskChecks.TryGetValue(t.TabId, out var c) && c;
                SyncStudentToTask(t, newName, checked_);
            }
            tab.ReloadFromDisk();
            ModernDialog.Alert("学生信息已保存");
            dlg.CloseAction?.Invoke();
        };
        ShowDialog("编辑学生", dlg, CreateStudentListView, 420, 480);
    }

    /// <summary>判断某任务的学生名单中是否包含指定学生姓名</summary>
    private static bool ContainsStudent(TaskTabViewModel task, string name)
    {
        var file = task.NameFilePath;
        if (!File.Exists(file)) return false;
        return File.ReadAllLines(file).Any(l => l.Trim() == name);
    }

    /// <summary>
    /// 将学生加入/移出指定任务的名单文件（name.txt），并刷新该任务
    /// </summary>
    private static void SyncStudentToTask(TaskTabViewModel task, string studentName, bool include)
    {
        var file = task.NameFilePath;
        var lines = File.Exists(file)
            ? File.ReadAllLines(file).Where(l => !string.IsNullOrWhiteSpace(l)).ToList()
            : new List<string>();
        var changed = false;
        if (include && !lines.Contains(studentName)) { lines.Add(studentName); changed = true; }
        else if (!include && lines.Contains(studentName)) { lines.Remove(studentName); changed = true; }
        if (changed)
        {
            File.WriteAllLines(file, lines);
            task.ReloadFromDisk();
        }
    }

    /// <summary>
    /// 收集所有任务（已打开的 + 后台的），用于学生编辑时勾选"需要打卡的任务"
    /// </summary>
    private List<TaskTabViewModel> GetAllTasks()
    {
        var all = new List<TaskTabViewModel>();
        all.AddRange(Tabs);
        all.AddRange(_backgroundTasks.Where(t => !all.Contains(t)));
        return all;
    }

    /// <summary>弹出文本输入对话框，返回输入内容（取消返回 null）</summary>
    private string? AskText(string title, string label)
    {
        string? result = null;
        var win = new Window
        {
            Title = title, Width = 360, Height = 190,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            ResizeMode = ResizeMode.NoResize, ShowInTaskbar = false
        };
        var panel = new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };
        panel.Children.Add(new System.Windows.Controls.TextBlock { Text = label, FontSize = 14, Margin = new Thickness(0, 0, 0, 10) });
        var txt = new System.Windows.Controls.TextBox { FontSize = 14, Padding = new Thickness(4) };
        panel.Children.Add(txt);
        var btnPanel = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment = System.Windows.HorizontalAlignment.Center, Margin = new Thickness(0, 15, 0, 0) };
        var okBtn = new System.Windows.Controls.Button { Content = "确定", Width = 80, Height = 30, Margin = new Thickness(5) };
        okBtn.Click += (_, _) => { result = txt.Text; win.Close(); };
        var cancelBtn = new System.Windows.Controls.Button { Content = "取消", Width = 80, Height = 30, Margin = new Thickness(5) };
        cancelBtn.Click += (_, _) => win.Close();
        btnPanel.Children.Add(okBtn); btnPanel.Children.Add(cancelBtn);
        panel.Children.Add(btnPanel);
        win.Content = panel;
        win.ShowDialog();
        return result;
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

    // ---- 检查更新 ----
    /// <summary>
    /// 检查客户端自身更新：直接查询 GitHub Releases 最新版本（不经过集控平台），比较后提示用户
    /// </summary>
    private async Task CheckForUpdateAsync()
    {
        StatusMessage = "正在检查更新...";
        string? latest = null;
        string? downloadUrl = null;

        // 1) 优先询问集控服务器（局域网，无需外网）
        try
        {
            var server = new ServerService();
            server.Initialize(Config.ServerIp, Config.ServerPort, Config.ServerPassword);
            var info = await server.GetClientUpdateAsync();
            if (info != null)
            {
                latest = info.Value.LatestVersion;
                downloadUrl = info.Value.DownloadUrl;
            }
        }
        catch { }

        // 2) 服务器不通时回退 GitHub 直查
        if (latest == null)
        {
            try
            {
                using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(8) };
                http.DefaultRequestHeaders.Add("User-Agent", "AgoraIn-Client");
                var resp = await http.GetAsync("https://api.github.com/repos/liuyuchen012/AgoraIn/releases/latest");
                if (resp.IsSuccessStatusCode)
                {
                    var json = await resp.Content.ReadFromJsonAsync<JsonElement>();
                    var tag = json.GetProperty("tag_name").GetString();
                    if (!string.IsNullOrEmpty(tag))
                        latest = tag!.StartsWith("v", StringComparison.OrdinalIgnoreCase) ? tag : "v" + tag;
                    foreach (var a in json.GetProperty("assets").EnumerateArray())
                        if ((a.GetProperty("name").GetString() ?? "") == "Client.win-x64.zip") { downloadUrl = a.GetProperty("browser_download_url").GetString(); break; }
                }
            }
            catch { }
        }

        if (latest == null)
        {
            ModernDialog.Alert("无法获取更新信息，请稍后重试或前往发布页查看。", "检查更新");
            StatusMessage = "检查更新失败";
            return;
        }

        if (IsNewerVersion(latest, AppConfig.Version))
        {
            var url = downloadUrl ?? "https://github.com/liuyuchen012/AgoraIn/releases";
            ModernDialog.UpdateAvailable(latest, AppConfig.Version, url, async () => await AutoUpdateAsync(url, latest));
            StatusMessage = $"发现新版本 {latest}";
        }
        else
        {
            ModernDialog.Alert($"当前已是最新版本 {AppConfig.Version}", "检查更新");
            StatusMessage = "已是最新版本";
        }
    }

    /// <summary>
    /// 自动更新：下载最新便携包 → 解压 → 生成 update.cmd（退出后替换并重启）
    /// 替换时跳过 config.json / workspace.json / 数据库 / data 目录，保留用户数据
    /// </summary>
    private async Task AutoUpdateAsync(string downloadUrl, string latest)
    {
        if (string.IsNullOrEmpty(downloadUrl) || !downloadUrl.StartsWith("https://"))
        {
            ModernDialog.Alert("更新地址不可用，请前往发布页手动下载。", "自动更新");
            return;
        }
        try
        {
            StatusMessage = $"正在下载更新包 {latest}...";
            var dir = Path.Combine(Path.GetTempPath(), "AgoraInUpdate", latest.TrimStart('v', 'V'));
            var pkgDir = Path.Combine(dir, "pkg");
            var cmdPath = Path.Combine(dir, "update.cmd");
            Directory.CreateDirectory(pkgDir);

            using (var http = new HttpClient { Timeout = TimeSpan.FromMinutes(15) })
            {
                using var resp = await http.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead);
                if (!resp.IsSuccessStatusCode)
                {
                    ModernDialog.Alert($"下载失败（HTTP {(int)resp.StatusCode}），请稍后重试或前往发布页手动下载。", "自动更新");
                    StatusMessage = "下载更新失败";
                    return;
                }
                var zipPath = Path.Combine(dir, "update.zip");
                await using (var fs = File.Create(zipPath))
                    await resp.Content.CopyToAsync(fs);
                if (Directory.Exists(pkgDir)) Directory.Delete(pkgDir, true);
                ZipFile.ExtractToDirectory(zipPath, pkgDir);
                File.Delete(zipPath);
            }

            // 校验更新包完整性
            if (!File.Exists(Path.Combine(pkgDir, "AgoraIn.exe")))
            {
                ModernDialog.Alert("更新包不完整，请前往发布页手动下载。", "自动更新");
                StatusMessage = "更新包校验失败";
                return;
            }

            var exeDir = AppDomain.CurrentDomain.BaseDirectory;
            var cmd = "@echo off\r\nchcp 65001 >nul\r\ntimeout /t 2 /nobreak >nul\r\n" +
                      "taskkill /im AgoraIn.exe /f >nul 2>&1\r\n" +
                      "robocopy \"" + pkgDir + "\" \"" + exeDir + "\" /E /XD \"" + exeDir + "\\data\" " +
                      "/XF config.json workspace.json *.db *.db-shm *.db-wal *.pdb >nul\r\n" +
                      "start \"\" \"" + exeDir + "\\AgoraIn.exe\"\r\n";
            File.WriteAllText(cmdPath, cmd);

            ModernDialog.Alert($"新版本 {latest} 已下载完成，程序即将重启并自动完成更新（保留你的配置与数据）。", "自动更新");
            Process.Start(new ProcessStartInfo("cmd.exe", "/c \"" + cmdPath + "\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            });
            Application.Current.Shutdown();
        }
        catch (Exception ex)
        {
            ModernDialog.Alert($"自动更新失败：{ex.Message}", "自动更新");
            StatusMessage = "自动更新失败";
        }
    }

    /// <summary>判断 latest 是否比 current 更新（语义化版本比较，支持 vX.Y.Z）</summary>
    private static bool IsNewerVersion(string latest, string current) => CompareVersion(latest, current) > 0;

    /// <summary>比较两个版本号，latest &gt; current 返回 1，相等返回 0，否则 -1</summary>
    private static int CompareVersion(string a, string b)
    {
        var pa = ParseVersion(a);
        var pb = ParseVersion(b);
        for (var i = 0; i < 3; i++)
        {
            var c = pa[i].CompareTo(pb[i]);
            if (c != 0) return c;
        }
        return 0;
    }

    /// <summary>将版本字符串解析为 [major, minor, patch] 三个整数</summary>
    private static int[] ParseVersion(string v)
    {
        var parts = v.Trim().TrimStart('v', 'V').Split('.');
        var nums = new int[3];
        for (var i = 0; i < 3; i++)
            if (i < parts.Length && int.TryParse(parts[i], out var n)) nums[i] = n;
        return nums;
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
    /// 构建学生列表管理对话框 UI 视图
    /// 列表模式：学生列表（支持多选）+ 添加/编辑/删除/批量删除/关闭
    /// 编辑模式：姓名输入框 + "需要进行打卡的任务"勾选列表 + 保存/取消
    /// </summary>
    private FrameworkElement CreateStudentListView(StudentListVm? vm)
    {
        if (vm == null) return new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };
        var panel = new System.Windows.Controls.StackPanel { Margin = new Thickness(15) };

        if (vm.IsEditMode)
        {
            // ---- 编辑模式 ----
            if (vm.SelectedStudent != null)
            {
                AddFieldRow(panel, "学生姓名:", vm.NewName, v => vm.NewName = v);
            }
            else
            {
                // 编辑模式必须有选中学生，否则回到列表模式
                panel.Children.Add(new System.Windows.Controls.TextBlock { Text = "请先选择学生再编辑", Margin = new Thickness(4) });
                AddSaveCancel(panel, vm);
                return panel;
            }

            panel.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = "需要进行打卡的任务:",
                FontSize = 13, FontWeight = System.Windows.FontWeights.SemiBold,
                Margin = new Thickness(4, 12, 4, 4)
            });
            var taskList = new System.Windows.Controls.StackPanel { Margin = new Thickness(4) };
            foreach (var t in vm.AllTasks)
            {
                var chk = new System.Windows.Controls.CheckBox
                {
                    Content = t.TabDisplayName,
                    IsChecked = vm.TaskChecks.TryGetValue(t.TabId, out var c) && c,
                    Margin = new Thickness(0, 4, 0, 4)
                };
                var tid = t.TabId;
                chk.Checked += (_, _) => vm.TaskChecks[tid] = true;
                chk.Unchecked += (_, _) => vm.TaskChecks[tid] = false;
                taskList.Children.Add(chk);
            }
            panel.Children.Add(taskList);
            AddSaveCancel(panel, vm);
            return panel;
        }

        // ---- 列表模式 ----
        panel.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = $"任务: {vm.Tab.TabDisplayName}  (共 {vm.Students.Count} 人)",
            FontSize = 13, FontWeight = System.Windows.FontWeights.SemiBold,
            Margin = new Thickness(4, 0, 4, 8)
        });
        var listBox = new System.Windows.Controls.ListBox
        {
            Height = 300, FontSize = 14,
            SelectionMode = System.Windows.Controls.SelectionMode.Extended,
            ItemsSource = vm.Students,
            DisplayMemberPath = nameof(StudentModel.Name),
            Margin = new Thickness(4)
        };
        listBox.SelectionChanged += (_, _) =>
        {
            vm.SelectedStudent = listBox.SelectedItem as StudentModel;
            vm.SelectedStudents.Clear();
            foreach (var s in listBox.SelectedItems.Cast<StudentModel>())
                if (s != null) vm.SelectedStudents.Add(s);
        };
        panel.Children.Add(listBox);
        panel.Children.Add(new System.Windows.Controls.TextBlock
        {
            Text = "提示: 按住 Ctrl 或 Shift 可多选，选中多名学生后可批量删除",
            FontSize = 11, Foreground = System.Windows.Media.Brushes.Gray,
            Margin = new Thickness(4, 2, 4, 0)
        });

        var btnPanel = new System.Windows.Controls.StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
            Margin = new Thickness(0, 12, 0, 0)
        };
        var add = new System.Windows.Controls.Button { Content = "添加", Width = 72, Height = 30, Margin = new Thickness(5) };
        add.Click += (_, _) => vm.AddAction?.Invoke();
        var edit = new System.Windows.Controls.Button { Content = "编辑", Width = 72, Height = 30, Margin = new Thickness(5) };
        edit.Click += (_, _) =>
        {
            if (vm.SelectedStudent == null) { ModernDialog.Alert("请先选中要编辑的学生"); return; }
            vm.EditAction?.Invoke();
        };
        var del = new System.Windows.Controls.Button { Content = "删除", Width = 72, Height = 30, Margin = new Thickness(5) };
        del.Click += (_, _) => vm.DeleteAction?.Invoke();
        var batchDel = new System.Windows.Controls.Button { Content = "批量删除", Width = 82, Height = 30, Margin = new Thickness(5) };
        batchDel.Click += (_, _) => vm.BatchDeleteAction?.Invoke();
        var close = new System.Windows.Controls.Button { Content = "关闭", Width = 72, Height = 30, Margin = new Thickness(5) };
        close.Click += (_, _) => vm.CloseAction?.Invoke();
        btnPanel.Children.Add(add); btnPanel.Children.Add(edit); btnPanel.Children.Add(del);
        btnPanel.Children.Add(batchDel); btnPanel.Children.Add(close);
        panel.Children.Add(btnPanel);
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

/// <summary>
/// 学生列表管理对话框的 ViewModel
/// 列表模式：管理当前任务学生（添加/编辑/删除/批量删除，支持多选）
/// 编辑模式：修改学生姓名 + 勾选该学生需要进行打卡的任务
/// </summary>
public class StudentListVm : IDialogVm
{
    /// <summary>当前操作的任务（活跃标签页）</summary>
    public TaskTabViewModel Tab { get; }
    /// <summary>当前任务的学生集合</summary>
    public ObservableCollection<StudentModel> Students => Tab.Students;
    /// <summary>所有任务（用于编辑时勾选"需要进行打卡的任务"）</summary>
    public List<TaskTabViewModel> AllTasks { get; }
    /// <summary>当前选中的学生（单选）</summary>
    public StudentModel? SelectedStudent { get; set; }
    /// <summary>当前选中的所有学生（多选，用于批量删除）</summary>
    public List<StudentModel> SelectedStudents { get; } = new();
    /// <summary>编辑模式下的新姓名</summary>
    public string NewName { get; set; } = "";
    /// <summary>是否为编辑模式（true 时显示姓名输入框与任务勾选列表）</summary>
    public bool IsEditMode { get; set; }
    /// <summary>任务 ID → 是否勾选（该学生需要在该任务打卡）</summary>
    public Dictionary<string, bool> TaskChecks { get; } = new();
    /// <summary>添加学生回调</summary>
    public Action? AddAction { get; set; }
    /// <summary>编辑学生回调</summary>
    public Action? EditAction { get; set; }
    /// <summary>删除学生回调</summary>
    public Action? DeleteAction { get; set; }
    /// <summary>批量删除学生回调</summary>
    public Action? BatchDeleteAction { get; set; }
    public Action? CloseAction { get; set; }
    public Action? SaveAction { get; set; }

    public StudentListVm(TaskTabViewModel tab, List<TaskTabViewModel> allTasks)
    {
        Tab = tab;
        AllTasks = allTasks;
    }
}
