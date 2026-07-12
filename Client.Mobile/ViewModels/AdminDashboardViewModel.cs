using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// 管理员仪表盘 ViewModel
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

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action<string>? NavigateRequested;

    public AdminDashboardViewModel(ApiService api)
    {
        _api = api;
        RefreshCommand = new Command(async () => await LoadDashboardAsync());
        NavigateToUsersCommand = new Command(() => NavigateRequested?.Invoke("users"));
        NavigateToTasksCommand = new Command(() => NavigateRequested?.Invoke("tasks"));
        NavigateToQRCodeCommand = new Command(() => NavigateRequested?.Invoke("qrcode"));
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

            // 设备列表
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

            // 活跃任务
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
