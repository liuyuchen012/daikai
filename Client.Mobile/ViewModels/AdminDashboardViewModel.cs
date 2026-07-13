using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// 管理员/教师仪表盘 ViewModel
/// </summary>
public class AdminDashboardViewModel : INotifyPropertyChanged
{
    private readonly ApiService _api;

    private bool _isLoading;
    public bool IsLoading { get => _isLoading; set { _isLoading = value; OnPropertyChanged(); } }

    private int _totalDevices;
    public int TotalDevices { get => _totalDevices; set { _totalDevices = value; OnPropertyChanged(); } }

    private int _onlineDevices;
    public int OnlineDevices { get => _onlineDevices; set { _onlineDevices = value; OnPropertyChanged(); } }

    private int _totalUsers;
    public int TotalUsers { get => _totalUsers; set { _totalUsers = value; OnPropertyChanged(); } }

    private int _todayCheckins;
    public int TodayCheckins { get => _todayCheckins; set { _todayCheckins = value; OnPropertyChanged(); } }

    private int _activeTasks;
    public int ActiveTasks { get => _activeTasks; set { _activeTasks = value; OnPropertyChanged(); } }

    public ObservableCollection<DeviceInfo> Devices { get; } = new();
    public ObservableCollection<ActiveTaskInfo> ActiveTaskList { get; } = new();

    public ICommand RefreshCommand { get; }
    public ICommand NavigateToUsersCommand { get; }
    public ICommand NavigateToTasksCommand { get; }
    public ICommand NavigateToQRCodeCommand { get; }
    public ICommand DeviceTappedCommand { get; }
    public ICommand RenameDeviceCommand { get; }
    public ICommand DeleteDeviceCommand { get; }
    public ICommand ViewDeviceQRCodeCommand { get; }
    public ICommand CreateTaskForDeviceCommand { get; }
    public ICommand DeleteTaskForDeviceCommand { get; }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action<string>? NavigateRequested;
    public event Action<string, string>? DeviceTapped; // uuid, name

    public AdminDashboardViewModel(ApiService api)
    {
        _api = api;
        RefreshCommand = new Command(async () =>
        {
            try { await LoadDashboardAsync(); }
            catch { /* 防止 async void 异常崩溃 */ }
        });
        NavigateToUsersCommand = new Command(() => NavigateRequested?.Invoke("adminusers"));
        NavigateToTasksCommand = new Command(() => NavigateRequested?.Invoke("tasks"));
        NavigateToQRCodeCommand = new Command(() => NavigateRequested?.Invoke("qrcode"));
        DeviceTappedCommand = new Command<DeviceInfo>(item =>
        {
            if (item != null)
                DeviceTapped?.Invoke(item.Uuid, item.Name);
        });
        RenameDeviceCommand = new Command<DeviceInfo>(async (item) =>
        {
            try { await RenameDeviceAsync(item); }
            catch { }
        });
        DeleteDeviceCommand = new Command<DeviceInfo>(async (item) =>
        {
            try { await DeleteDeviceAsync(item); }
            catch { }
        });
        ViewDeviceQRCodeCommand = new Command<DeviceInfo>(async (item) =>
        {
            try { await ViewDeviceTasksAsync(item); }
            catch { }
        });
        CreateTaskForDeviceCommand = new Command<DeviceInfo>(async (item) =>
        {
            try { await CreateTaskForDeviceAsync(item); }
            catch { }
        });
        DeleteTaskForDeviceCommand = new Command<DeviceInfo>(async (item) =>
        {
            try { await DeleteTaskForDeviceAsync(item); }
            catch { }
        });
    }

    public async Task LoadDashboardAsync()
    {
        IsLoading = true;
        try
        {
            var result = await _api.GetAsync("/api/mobile/dashboard");
            if (ApiService.GetError(result) != null) return;

            if (result.TryGetProperty("summary", out var summary))
            {
                TotalDevices = GetInt(summary, "total_devices");
                OnlineDevices = GetInt(summary, "online_devices");
                TotalUsers = GetInt(summary, "total_users");
                TodayCheckins = GetInt(summary, "today_checkins");
                ActiveTasks = GetInt(summary, "active_signin_tasks");
            }

            Devices.Clear();
            if (result.TryGetProperty("devices", out var devices) && devices.ValueKind == JsonValueKind.Array)
            {
                foreach (var d in devices.EnumerateArray())
                {
                    Devices.Add(new DeviceInfo
                    {
                        Name = ApiService.GetString(d, "name") ?? "未知",
                        Uuid = ApiService.GetString(d, "uuid") ?? "",
                        Online = d.TryGetProperty("online", out var on) && on.GetBoolean(),
                        TaskCount = GetInt(d, "task_count")
                    });
                }
            }

            ActiveTaskList.Clear();
            if (result.TryGetProperty("active_signin_tasks", out var tasks) && tasks.ValueKind == JsonValueKind.Array)
            {
                foreach (var t in tasks.EnumerateArray())
                {
                    ActiveTaskList.Add(new ActiveTaskInfo
                    {
                        ShortCode = ApiService.GetString(t, "short_code") ?? "",
                        Subject = ApiService.GetString(t, "subject") ?? "未知",
                        Classroom = ApiService.GetString(t, "classroom") ?? "",
                        StudentCount = GetInt(t, "student_count"),
                        SignedCount = GetInt(t, "signed_count")
                    });
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Dashboard load error: {ex.Message}");
        }
        finally
        {
            IsLoading = false;
        }
    }

    private async Task RenameDeviceAsync(DeviceInfo? item)
    {
        if (item == null) return;
        var newName = await Shell.Current.DisplayPromptAsync("重命名设备",
            $"请输入设备「{item.Name}」的新名称：", "确定", "取消",
            placeholder: "新设备名称", initialValue: item.Name);
        if (string.IsNullOrWhiteSpace(newName) || newName == item.Name) return;

        IsLoading = true;
        try
        {
            var result = await _api.PutAsync($"/api/mobile/devices/{Uri.EscapeDataString(item.Uuid)}/rename",
                new { name = newName.Trim() });
            var error = ApiService.GetError(result);
            if (error != null)
            {
                await Shell.Current.DisplayAlertAsync("重命名失败", error, "确定");
            }
            else
            {
                item.Name = newName.Trim();
                await LoadDashboardAsync(); // 刷新列表
            }
        }
        catch (Exception ex)
        {
            await Shell.Current.DisplayAlertAsync("重命名失败", $"网络错误: {ex.Message}", "确定");
        }
        finally { IsLoading = false; }
    }

    private async Task DeleteDeviceAsync(DeviceInfo? item)
    {
        if (item == null) return;
        if (!await Shell.Current.DisplayAlertAsync("确认删除",
            $"确定要删除设备「{item.Name}」吗？\n所有绑定的任务和签到数据也将被删除，此操作不可撤销！",
            "删除", "取消"))
            return;

        IsLoading = true;
        try
        {
            var result = await _api.DeleteAsync($"/api/mobile/devices/{Uri.EscapeDataString(item.Uuid)}");
            var error = ApiService.GetError(result);
            if (error != null)
                await Shell.Current.DisplayAlertAsync("删除失败", error, "确定");
            else
                await LoadDashboardAsync();
        }
        catch (Exception ex)
        {
            await Shell.Current.DisplayAlertAsync("删除失败", $"网络错误: {ex.Message}", "确定");
        }
        finally { IsLoading = false; }
    }

    private async Task ViewDeviceTasksAsync(DeviceInfo? item)
    {
        if (item == null) return;
        // 导航到考勤详情页面，查看该设备下的所有任务
        DeviceTapped?.Invoke(item.Uuid, item.Name);
    }

    private async Task CreateTaskForDeviceAsync(DeviceInfo? item)
    {
        if (item == null) return;
        var subject = await Shell.Current.DisplayPromptAsync("创建普通任务",
            $"为设备「{item.Name}」创建新任务\n请输入科目名称：", "下一步", "取消",
            placeholder: "如：语文");
        if (string.IsNullOrWhiteSpace(subject)) return;

        var classroom = await Shell.Current.DisplayPromptAsync("创建普通任务",
            $"科目「{subject.Trim()}」\n请输入教室（可选）：", "创建", "取消",
            placeholder: "如：三年(1)班");
        if (classroom == null) return; // 取消

        IsLoading = true;
        try
        {
            var result = await _api.PostAsync($"/api/mobile/devices/{Uri.EscapeDataString(item.Uuid)}/tasks",
                new { subject = subject.Trim(), classroom = (classroom.Trim() ?? ""), task_name = subject.Trim() });
            var error = ApiService.GetError(result);
            if (error != null)
                await Shell.Current.DisplayAlertAsync("创建失败", error, "确定");
            else
            {
                await Shell.Current.DisplayAlertAsync("创建成功", $"任务「{subject.Trim()}」已推送至设备「{item.Name}」", "确定");
                await LoadDashboardAsync();
            }
        }
        catch (Exception ex)
        {
            await Shell.Current.DisplayAlertAsync("创建失败", $"网络错误: {ex.Message}", "确定");
        }
        finally { IsLoading = false; }
    }

    private async Task DeleteTaskForDeviceAsync(DeviceInfo? item)
    {
        if (item == null) return;
        try
        {
            var statusResult = await _api.GetAsync("/api/status");
            if (statusResult.ValueKind != JsonValueKind.Array) { await LoadDashboardAsync(); return; }
            var statusList = new List<JsonElement>();
            foreach (var e in statusResult.EnumerateArray()) statusList.Add(e);

            var deviceStatus = statusList.FirstOrDefault(d =>
                ApiService.GetString(d, "uuid") == item.Uuid);
            if (deviceStatus.ValueKind != JsonValueKind.Object) { await LoadDashboardAsync(); return; }

            var tasks = new List<string>();
            if (deviceStatus.TryGetProperty("tasks", out var taskArr) && taskArr.ValueKind == JsonValueKind.Array)
                tasks = taskArr.EnumerateArray().Select(t => t.GetString() ?? "").Where(s => !string.IsNullOrEmpty(s)).ToList();

            if (tasks.Count == 0)
            {
                await Shell.Current.DisplayAlertAsync("提示", $"设备「{item.Name}」上暂无任务", "确定");
                return;
            }

            string targetTaskId;
            if (tasks.Count == 1)
            {
                targetTaskId = tasks[0];
                if (!await Shell.Current.DisplayAlertAsync("确认删除",
                    $"确定要删除设备「{item.Name}」上的唯一任务吗？\n相关打卡数据也将被删除。", "删除", "取消"))
                    return;
            }
            else
            {
                var taskStr = await Shell.Current.DisplayPromptAsync("删除任务",
                    $"设备「{item.Name}」上有 {tasks.Count} 个任务\n输入要删除的任务索引 (1-{tasks.Count})：", "删除", "取消",
                    placeholder: "1",
                    keyboard: Keyboard.Numeric);
                if (string.IsNullOrWhiteSpace(taskStr) || !int.TryParse(taskStr, out var idx) || idx < 1 || idx > tasks.Count)
                    return;
                targetTaskId = tasks[idx - 1];
            }

            IsLoading = true;
            var result = await _api.DeleteAsync($"/api/mobile/devices/{Uri.EscapeDataString(item.Uuid)}/tasks/{Uri.EscapeDataString(targetTaskId)}");
            var error = ApiService.GetError(result);
            if (error != null)
                await Shell.Current.DisplayAlertAsync("删除失败", error, "确定");
            else
            {
                await Shell.Current.DisplayAlertAsync("删除成功", $"已从设备「{item.Name}」删除任务", "确定");
                await LoadDashboardAsync();
            }
        }
        catch (Exception ex)
        {
            await Shell.Current.DisplayAlertAsync("操作失败", $"网络错误: {ex.Message}", "确定");
        }
        finally { IsLoading = false; }
    }

    private static int GetInt(JsonElement json, string key) =>
        json.TryGetProperty(key, out var val) && val.TryGetInt32(out var i) ? i : 0;

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public class DeviceInfo
{
    public string Name { get; set; } = "";
    public string Uuid { get; set; } = "";
    public bool Online { get; set; }
    public int TaskCount { get; set; }
    public string StatusText => Online ? "在线" : "离线";
    public Color StatusColor => Online ? Color.FromArgb("#34a853") : Color.FromArgb("#888888");
}

public class ActiveTaskInfo
{
    public string ShortCode { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Classroom { get; set; } = "";
    public int StudentCount { get; set; }
    public int SignedCount { get; set; }
    public string ProgressText => $"{SignedCount}/{StudentCount}";
    public double Progress => StudentCount > 0 ? (double)SignedCount / StudentCount : 0;
}
