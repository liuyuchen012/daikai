using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using CheckIn.Client.Models;
using CheckIn.Client.Services;

namespace CheckIn.Client.ViewModels;

/// <summary>
/// 远程打卡服务器控制 ViewModel：登录、仪表盘、设备、任务、考勤、用户管理
/// </summary>
public class ServerControlViewModel : INotifyPropertyChanged
{
    private readonly RemoteControlService _api;

    // ---- 登录状态 ----
    private bool _isLoggedIn;
    public bool IsLoggedIn
    {
        get => _isLoggedIn;
        set { _isLoggedIn = value; OnPropertyChanged(); OnPropertyChanged(nameof(IsNotLoggedIn)); }
    }
    public bool IsNotLoggedIn => !_isLoggedIn;

    private string _serverUrl = "http://192.168.1.100:5000";
    public string ServerUrl
    {
        get => _serverUrl;
        set { _serverUrl = value; OnPropertyChanged(); }
    }

    private string _username = "";
    public string Username
    {
        get => _username;
        set { _username = value; OnPropertyChanged(); }
    }

    private string _password = "";
    public string Password
    {
        get => _password;
        set { _password = value; OnPropertyChanged(); }
    }

    private string _loginMessage = "";
    public string LoginMessage
    {
        get => _loginMessage;
        set { _loginMessage = value; OnPropertyChanged(); }
    }

    private string _userInfo = "";
    public string UserInfo
    {
        get => _userInfo;
        set { _userInfo = value; OnPropertyChanged(); }
    }

    private bool _isAdmin;
    public bool IsAdmin
    {
        get => _isAdmin;
        set { _isAdmin = value; OnPropertyChanged(); }
    }

    private string _statusMessage = "";
    public string StatusMessage
    {
        get => _statusMessage;
        set { _statusMessage = value; OnPropertyChanged(); }
    }

    // ---- 仪表盘 ----
    private DashboardSummary? _summary;
    public DashboardSummary? Summary
    {
        get => _summary;
        set { _summary = value; OnPropertyChanged(); }
    }

    public ObservableCollection<DashboardDevice> Devices { get; } = new();
    public ObservableCollection<DashboardTask> ActiveTasks { get; } = new();

    // ---- 设备 ----
    public ObservableCollection<DeviceItem> DeviceItems { get; } = new();

    // ---- 任务 ----
    public ObservableCollection<RemoteTask> TaskItems { get; } = new();

    // ---- 考勤 ----
    public ObservableCollection<AttendanceTask> AttendanceItems { get; } = new();

    // ---- 历史 ----
    public ObservableCollection<HistoryTask> HistoryItems { get; } = new();

    // ---- 用户 ----
    public ObservableCollection<RemoteUserItem> UserItems { get; } = new();

    public ServerControlViewModel()
    {
        _api = new RemoteControlService();
    }

    // ============ 登录 / 登出 ============

    public async Task<bool> LoginAsync()
    {
        try
        {
            LoginMessage = "登录中…";
            _api.SetBaseUrl(ServerUrl);
            var user = await _api.LoginAsync(Username, Password);
            IsLoggedIn = true;
            IsAdmin = user.Role == "admin";
            UserInfo = $"{user.DisplayName}（{user.Username}）· {user.Role}";
            LoginMessage = "";
            return true;
        }
        catch (Exception ex)
        {
            LoginMessage = ex.Message;
            return false;
        }
    }

    public void Logout()
    {
        _api.Logout();
        IsLoggedIn = false;
        IsAdmin = false;
        UserInfo = "";
        Password = "";
        ClearAll();
    }

    private void ClearAll()
    {
        Devices.Clear();
        ActiveTasks.Clear();
        DeviceItems.Clear();
        TaskItems.Clear();
        AttendanceItems.Clear();
        HistoryItems.Clear();
        UserItems.Clear();
        Summary = null;
    }

    // ============ 数据加载 ============

    public async Task<bool> LoadDashboardAsync()
    {
        try
        {
            var data = await _api.GetDashboardAsync();
            Summary = data.Summary;
            Devices.Clear();
            foreach (var d in data.Devices) Devices.Add(d);
            ActiveTasks.Clear();
            foreach (var t in data.ActiveSignInTasks) ActiveTasks.Add(t);
            StatusMessage = "仪表盘已刷新";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    public async Task<bool> LoadDevicesAsync()
    {
        try
        {
            var list = await _api.GetDevicesAsync();
            DeviceItems.Clear();
            foreach (var d in list) DeviceItems.Add(d);
            StatusMessage = $"设备 {list.Count} 台";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    public async Task<bool> LoadTasksAsync()
    {
        try
        {
            var list = await _api.GetTasksAsync();
            TaskItems.Clear();
            foreach (var t in list) TaskItems.Add(t);
            StatusMessage = $"任务 {list.Count} 个";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    public async Task<bool> LoadAttendanceAsync()
    {
        try
        {
            var list = await _api.GetAttendanceAsync();
            AttendanceItems.Clear();
            foreach (var t in list) AttendanceItems.Add(t);
            StatusMessage = $"考勤记录 {list.Count} 条";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    public async Task<bool> LoadHistoryAsync()
    {
        try
        {
            var data = await _api.GetHistoryAsync();
            HistoryItems.Clear();
            foreach (var h in data.History) HistoryItems.Add(h);
            StatusMessage = $"历史任务 {data.TotalTasks} 个，签到 {data.TotalCheckins} 人次";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    public async Task<bool> LoadUsersAsync()
    {
        if (!IsAdmin)
        {
            StatusMessage = "仅管理员可查看用户";
            return false;
        }
        try
        {
            var list = await _api.GetUsersAsync();
            UserItems.Clear();
            foreach (var u in list) UserItems.Add(u);
            StatusMessage = $"用户 {list.Count} 个";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    // ============ 设备操作 ============

    public async Task<bool> RenameDeviceAsync(DeviceItem device, string newName)
    {
        try
        {
            await _api.RenameDeviceAsync(device.Uuid, newName);
            device.Name = newName;
            StatusMessage = "设备已重命名";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    public async Task<bool> DeleteDeviceAsync(DeviceItem device)
    {
        try
        {
            await _api.DeleteDeviceAsync(device.Uuid);
            DeviceItems.Remove(device);
            StatusMessage = "设备已删除";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    // ============ 任务操作 ============

    public async Task<bool> CloseTaskAsync(RemoteTask task)
    {
        try
        {
            await _api.CloseTaskAsync(task.Id);
            task.Status = "closed";
            OnPropertyChanged(nameof(TaskItems));
            StatusMessage = "任务已关闭";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    public async Task<bool> DeleteTaskAsync(RemoteTask task)
    {
        try
        {
            await _api.DeleteTaskAsync(task.Id);
            TaskItems.Remove(task);
            StatusMessage = "任务已删除";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    public async Task<bool> RenameTaskAsync(RemoteTask task, string newName)
    {
        try
        {
            await _api.RenameTaskAsync(task.Id, newName);
            task.Subject = newName;
            OnPropertyChanged(nameof(TaskItems));
            StatusMessage = "任务已重命名";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    // ============ 用户操作 ============

    public async Task<bool> CreateUserAsync(string username, string password, string role, string displayName)
    {
        try
        {
            await _api.CreateUserAsync(username, password, role, displayName);
            StatusMessage = "用户已创建";
            await LoadUsersAsync();
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    public async Task<bool> DeleteUserAsync(RemoteUserItem user)
    {
        try
        {
            await _api.DeleteUserAsync(user.Id);
            UserItems.Remove(user);
            StatusMessage = "用户已删除";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    public async Task<bool> ToggleUserActiveAsync(RemoteUserItem user)
    {
        try
        {
            await _api.UpdateUserAsync(user.Id, isActive: !user.IsActive);
            user.IsActive = !user.IsActive;
            OnPropertyChanged(nameof(UserItems));
            StatusMessage = user.IsActive ? "用户已启用" : "用户已禁用";
            return true;
        }
        catch (Exception ex)
        {
            StatusMessage = ex.Message;
            return false;
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
