using System.Collections.ObjectModel;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Mobile.Models;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// Main application ViewModel. Manages tasks, tabs, global config, and server connection.
/// Adapted from WPF MainViewModel for MAUI mobile.
/// </summary>
public class MainViewModel : BaseViewModel
{
    public AppConfig Config { get; } = new();

    private readonly string _baseDir;
    private readonly string _configFile;
    private readonly string _workspaceFile;

    private readonly ServerService _serverService;

    /// <summary>
    /// Collection of all task tabs (both open and background).
    /// </summary>
    public ObservableCollection<TaskTabViewModel> Tabs { get; } = new();

    /// <summary>
    /// Background tasks (not currently displayed but still syncing).
    /// </summary>
    private List<TaskTabViewModel> _backgroundTasks = new();

    /// <summary>
    /// All tasks flattened for display in task list.
    /// </summary>
    public ObservableCollection<TaskTabViewModel> AllTasks { get; } = new();

    private TaskTabViewModel? _activeTab;
    public TaskTabViewModel? ActiveTab
    {
        get => _activeTab;
        set { SetProperty(ref _activeTab, value); OnPropertyChanged(nameof(HasActiveTab)); }
    }

    public bool HasActiveTab => ActiveTab != null;

    public ICommand AddTabCommand { get; }
    public ICommand DeleteTabCommand { get; }
    public ICommand ShowSettingsCommand { get; }
    public ICommand CreateSignInCommand { get; }

    public MainViewModel()
    {
        _baseDir = FileSystem.AppDataDirectory;
        _configFile = Path.Combine(_baseDir, "config.json");
        _workspaceFile = Path.Combine(_baseDir, "workspace.json");

        _serverService = new ServerService();

        LoadGlobalConfig();

        AddTabCommand = new RelayCommand(_ => AddTab());
        DeleteTabCommand = new RelayCommand(DeleteTab);
        ShowSettingsCommand = new RelayCommand(async _ => await Shell.Current.GoToAsync("settings"));
        CreateSignInCommand = new RelayCommand(async _ => await Shell.Current.GoToAsync("createsignin"));

        LoadWorkspace();
        InitializeBackgroundTasks();
    }

    // ---- Task Management ----
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
        if (string.IsNullOrEmpty(tab.Config.Name)) tab.Config.Name = $"Task {Tabs.Count + 1}";
        tab.SaveConfig();
        if (isSignIn && !string.IsNullOrEmpty(signInTaskId))
            tab.UpdateGlobalServerConfig(Config.ServerIp, Config.ServerPort, Config.ServerPassword, AppConfig.Version);

        Tabs.Add(tab);
        ActiveTab = tab;
        RefreshAllTasks();
        SaveWorkspace();
    }

    private void DeleteTab(object? param)
    {
        if (param is not TaskTabViewModel tab) return;

        if (Tabs.Contains(tab))
            Tabs.Remove(tab);

        _backgroundTasks.Remove(tab);
        tab.Dispose();

        var tabDir = Path.Combine(_baseDir, "data", "tabs", tab.TabId);
        if (Directory.Exists(tabDir))
            Directory.Delete(tabDir, recursive: true);

        if (ActiveTab == tab)
            ActiveTab = Tabs.FirstOrDefault() ?? _backgroundTasks.FirstOrDefault();

        RefreshAllTasks();
        SaveWorkspace();
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
        tab.Title = tab.TabDisplayName;
        RefreshAllTasks();
        SaveWorkspace();
    }

    private void RefreshAllTasks()
    {
        AllTasks.Clear();
        foreach (var t in Tabs.OrderBy(t => t.Config.Name))
            AllTasks.Add(t);
        foreach (var t in _backgroundTasks.OrderBy(t => t.Config.Name))
            if (!AllTasks.Contains(t))
                AllTasks.Add(t);
    }

    private void InitializeBackgroundTasks()
    {
        var tabsDir = Path.Combine(_baseDir, "data", "tabs");
        if (!Directory.Exists(tabsDir)) return;

        foreach (var tabDir in Directory.GetDirectories(tabsDir))
        {
            var tabId = Path.GetFileName(tabDir);
            if (Tabs.Any(t => t.TabId == tabId)) continue;

            var bgTab = TaskTabViewModel.CreateBackgroundInstance(tabId, _baseDir, Config);
            _backgroundTasks.Add(bgTab);
        }
    }

    // ---- Workspace Persistence ----
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
                    RefreshAllTasks();
                    return;
                }
            }
            catch { }
        }

        // First launch: create a default tab
        AddTab("Default Task", "Math");
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

    // ---- Global Config ----
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

    public void UpdateServerConfig(string ip, int port, string password)
    {
        Config.ServerIp = ip;
        Config.ServerPort = port;
        Config.ServerPassword = password;
        SaveGlobalConfig();

        foreach (var tab in Tabs)
            tab.UpdateGlobalServerConfig(ip, port, password, AppConfig.Version);
        foreach (var bgTab in _backgroundTasks)
            bgTab.UpdateGlobalServerConfig(ip, port, password, AppConfig.Version);
    }

    // ---- Sign In Task ----
    public async Task<(string shortCode, string taskId)?> CreateRemoteSignInAsync(string password, string classroom, string subject, List<string> students)
    {
        var server = new ServerService();
        server.Initialize(Config.ServerIp, Config.ServerPort, Config.ServerPassword);
        await server.RegisterAsync("AgoraIn Mobile");
        return await server.CreateSignInAsync(password, classroom, subject, students);
    }

    // ---- Helpers ----
    public static string HashPassword(string pwd)
    {
        using var sha = System.Security.Cryptography.SHA256.Create();
        return Convert.ToHexString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(pwd)));
    }
}

public class WorkspaceConfig
{
    public List<TabInfo> Tabs { get; set; } = new();
    public string? ActiveTabId { get; set; }
}
