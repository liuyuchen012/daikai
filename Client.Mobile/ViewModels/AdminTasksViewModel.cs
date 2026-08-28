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
    public ICommand RenameTaskCommand { get; }
    public ICommand ViewQRCodeCommand { get; }

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action<string, string>? ViewAttendanceRequested; // shortCode, subject

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
        RenameTaskCommand = new Command<TaskItem>(async (item) =>
        {
            try { await RenameTaskAsync(item); }
            catch { }
        });
        ViewQRCodeCommand = new Command<TaskItem>(async (item) =>
        {
            try { await ViewTaskQRCodeAsync(item); }
            catch { }
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
            if (!await Shell.Current.DisplayAlertAsync("确认删除", "确定要删除该签到任务吗？", "删除", "取消"))
                return;
            await _api.DeleteAsync($"/api/mobile/tasks/{id}");
            await LoadTasksAsync();
        }
        catch { }
    }

    private async Task RenameTaskAsync(TaskItem? item)
    {
        if (item == null) return;
        var newName = await Shell.Current.DisplayPromptAsync("重命名任务",
            $"请输入任务的新名称：", "确定", "取消",
            placeholder: "新任务名称", initialValue: item.Subject);
        if (string.IsNullOrWhiteSpace(newName) || newName == item.Subject) return;

        IsLoading = true;
        try
        {
            var result = await _api.PutAsync($"/api/mobile/tasks/{item.Id}/rename",
                new { name = newName.Trim() });
            var error = ApiService.GetError(result);
            if (error != null)
                await Shell.Current.DisplayAlertAsync("重命名失败", error, "确定");
            else
                await LoadTasksAsync();
        }
        catch (Exception ex)
        {
            await Shell.Current.DisplayAlertAsync("重命名失败", $"网络错误: {ex.Message}", "确定");
        }
        finally { IsLoading = false; }
    }

    private async Task ViewTaskQRCodeAsync(TaskItem? item)
    {
        if (item == null) return;
        try
        {
            var result = await _api.GetAsync($"/api/mobile/tasks/{item.Id}/qrcode");
            var error = ApiService.GetError(result);
            if (error != null)
            {
                await Shell.Current.DisplayAlertAsync("错误", error, "确定");
                return;
            }

            var shortCode = ApiService.GetString(result, "short_code") ?? "";
            var baseUrl = Preferences.Get("server_url", "http://localhost:5250");
            var url = $"{baseUrl}/s/{shortCode}";
            var subject = ApiService.GetString(result, "subject") ?? "";
            var status = ApiService.GetString(result, "status") ?? "";
            var signedCount = GetInt(result, "signed_count");
            var studentCount = GetInt(result, "student_count");

            var action = await Shell.Current.DisplayAlertAsync(
                "签到二维码",
                $"任务: {subject}\n状态: {(status == "active" ? "进行中" : "已关闭")}\n签到: {signedCount}/{studentCount}",
                "复制链接", "关闭");
            if (action)
            {
                await Clipboard.Default.SetTextAsync(url);
                await Shell.Current.DisplayAlertAsync("已复制", "签到链接已复制到剪贴板", "确定");
            }
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
