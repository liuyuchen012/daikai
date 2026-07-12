using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows.Input;
using CheckIn.Client.Mobile.Services;

namespace CheckIn.Client.Mobile.ViewModels;

/// <summary>
/// 管理员任务管理 ViewModel
/// </summary>
public class AdminTasksViewModel : INotifyPropertyChanged
{
    private readonly ApiService _api;

    private bool _isLoading;
    public bool IsLoading { get => _isLoading; set { _isLoading = value; OnPropertyChanged(); } }

    public ObservableCollection<TaskItem> Tasks { get; } = new();
    public ICommand RefreshCommand { get; }
    public ICommand CloseTaskCommand { get; }
    public ICommand DeleteTaskCommand { get; }
    public ICommand ViewAttendanceCommand { get; }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action<string, string>? ViewAttendanceRequested; // taskId, subject

    public AdminTasksViewModel(ApiService api)
    {
        _api = api;
        RefreshCommand = new Command(async () =>
        {
            try { await LoadTasksAsync(); }
            catch { /* 防止 async void 异常崩溃 */ }
        });
        CloseTaskCommand = new Command<int>(async (id) =>
        {
            try { await CloseTaskAsync(id); }
            catch { /* 防止 async void 异常崩溃 */ }
        });
        DeleteTaskCommand = new Command<int>(async (id) =>
        {
            try { await DeleteTaskAsync(id); }
            catch { /* 防止 async void 异常崩溃 */ }
        });
        ViewAttendanceCommand = new Command<TaskItem>(item =>
        {
            if (item != null)
                ViewAttendanceRequested?.Invoke(item.ShortCode, item.Subject);
        });
    }

    public async Task LoadTasksAsync()
    {
        IsLoading = true;
        try
        {
            var result = await _api.GetAsync("/api/mobile/tasks");
            if (ApiService.GetError(result) != null) return;

            Tasks.Clear();
            if (result.TryGetProperty("tasks", out var tasks) && tasks.ValueKind == JsonValueKind.Array)
            {
                foreach (var t in tasks.EnumerateArray())
                {
                    Tasks.Add(new TaskItem
                    {
                        Id = GetInt(t, "id"),
                        ShortCode = ApiService.GetString(t, "short_code") ?? "",
                        Subject = ApiService.GetString(t, "subject") ?? "未知",
                        Classroom = ApiService.GetString(t, "classroom") ?? "",
                        Status = ApiService.GetString(t, "status") ?? "active",
                        StudentCount = GetInt(t, "student_count"),
                        SignedCount = GetInt(t, "signed_count"),
                        CreatedAt = ApiService.GetString(t, "created_at") ?? ""
                    });
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Tasks load error: {ex.Message}");
        }
        finally
        {
            IsLoading = false;
        }
    }

    private async Task CloseTaskAsync(int id)
    {
        try
        {
            await _api.PostAsync($"/api/mobile/tasks/{id}/close");
            await LoadTasksAsync();
        }
        catch { }
    }

    private async Task DeleteTaskAsync(int id)
    {
        try
        {
            await _api.DeleteAsync($"/api/mobile/tasks/{id}");
            await LoadTasksAsync();
        }
        catch { }
    }

    private static int GetInt(JsonElement json, string key) =>
        json.TryGetProperty(key, out var val) && val.TryGetInt32(out var i) ? i : 0;

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

public class TaskItem
{
    public int Id { get; set; }
    public string ShortCode { get; set; } = "";
    public string Subject { get; set; } = "";
    public string Classroom { get; set; } = "";
    public string Status { get; set; } = "active";
    public int StudentCount { get; set; }
    public int SignedCount { get; set; }
    public string CreatedAt { get; set; } = "";
    public string ProgressText => $"{SignedCount}/{StudentCount}";
    public string StatusText => Status == "active" ? "进行中" : "已关闭";
    public Color StatusColor => Status == "active" ? Color.FromArgb("#34a853") : Color.FromArgb("#888888");
    public double Progress => StudentCount > 0 ? (double)SignedCount / StudentCount : 0;
}
