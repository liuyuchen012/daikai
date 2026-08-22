using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;
using CheckIn.Client.Models;

namespace CheckIn.Client.Services;

/// <summary>
/// 远程打卡服务器控制服务：登录、仪表盘、设备、任务、考勤、用户管理
/// 使用 Bearer Token 认证，对应服务器 /api/auth/* 与 /api/mobile/* 接口
/// </summary>
public class RemoteControlService
{
    private readonly HttpClient _http;
    private string _baseUrl = "";
    private string _token = "";
    private static readonly JsonSerializerOptions _jsonOptions = new() { PropertyNameCaseInsensitive = true };

    public string BaseUrl => _baseUrl;
    public string? Token => string.IsNullOrEmpty(_token) ? null : _token;
    public bool IsLoggedIn => !string.IsNullOrEmpty(_token);

    public RemoteControlService()
    {
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(10) };
    }

    /// <summary>设置服务器地址，格式如 http://192.168.1.100:5000</summary>
    public void SetBaseUrl(string baseUrl)
    {
        _baseUrl = baseUrl.TrimEnd('/');
    }

    private HttpRequestMessage MakeRequest(HttpMethod method, string path, object? body = null)
    {
        var req = new HttpRequestMessage(method, $"{_baseUrl}{path}");
        if (!string.IsNullOrEmpty(_token))
            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _token);
        if (body != null)
            req.Content = JsonContent.Create(body);
        return req;
    }

    private async Task<JsonElement?> SendAsync(HttpRequestMessage req)
    {
        var res = await _http.SendAsync(req);
        if (!res.IsSuccessStatusCode)
        {
            string? errMsg = null;
            try
            {
                var err = await res.Content.ReadFromJsonAsync<JsonElement>();
                if (err.TryGetProperty("error", out var e))
                    errMsg = e.GetString();
            }
            catch { }
            throw new InvalidOperationException(errMsg ?? $"请求失败 ({res.StatusCode})");
        }
        return await res.Content.ReadFromJsonAsync<JsonElement>();
    }

    // ============ 认证 ============

    /// <summary>登录，成功保存 Token</summary>
    public async Task<RemoteUser> LoginAsync(string username, string password)
    {
        var res = await _http.PostAsJsonAsync($"{_baseUrl}/api/auth/login", new { username, password });
        var json = await res.Content.ReadFromJsonAsync<JsonElement>();
        if (!res.IsSuccessStatusCode)
        {
            var err = json.TryGetProperty("error", out var e) ? e.GetString() : "登录失败";
            throw new InvalidOperationException(err ?? "登录失败");
        }
        _token = json.GetProperty("token").GetString() ?? "";
        return JsonSerializer.Deserialize<RemoteUser>(json.GetProperty("user").GetRawText(), _jsonOptions) ?? new RemoteUser();
    }

    public void Logout() => _token = "";

    // ============ 仪表盘 ============

    public async Task<DashboardResponse> GetDashboardAsync()
    {
        using var req = MakeRequest(HttpMethod.Get, "/api/mobile/dashboard");
        var json = await SendAsync(req);
        return JsonSerializer.Deserialize<DashboardResponse>(json?.GetRawText() ?? "{}", _jsonOptions) ?? new DashboardResponse();
    }

    // ============ 设备 ============

    public async Task<List<DeviceItem>> GetDevicesAsync()
    {
        using var req = MakeRequest(HttpMethod.Get, "/api/mobile/devices");
        var json = await SendAsync(req);
        return JsonSerializer.Deserialize<DeviceResponse>(json?.GetRawText() ?? "{}", _jsonOptions)?.Devices ?? new();
    }

    public async Task RenameDeviceAsync(string uuid, string name)
    {
        using var req = MakeRequest(HttpMethod.Put, $"/api/mobile/devices/{uuid}/rename", new { name });
        await SendAsync(req);
    }

    public async Task DeleteDeviceAsync(string uuid)
    {
        using var req = MakeRequest(HttpMethod.Delete, $"/api/mobile/devices/{uuid}");
        await SendAsync(req);
    }

    // ============ 任务 ============

    public async Task<List<RemoteTask>> GetTasksAsync()
    {
        using var req = MakeRequest(HttpMethod.Get, "/api/mobile/tasks");
        var json = await SendAsync(req);
        return JsonSerializer.Deserialize<TaskListResponse>(json?.GetRawText() ?? "{}", _jsonOptions)?.Tasks ?? new();
    }

    public async Task CloseTaskAsync(int id)
    {
        using var req = MakeRequest(HttpMethod.Post, $"/api/mobile/tasks/{id}/close");
        await SendAsync(req);
    }

    public async Task DeleteTaskAsync(int id)
    {
        using var req = MakeRequest(HttpMethod.Delete, $"/api/mobile/tasks/{id}");
        await SendAsync(req);
    }

    public async Task RenameTaskAsync(int id, string name)
    {
        using var req = MakeRequest(HttpMethod.Put, $"/api/mobile/tasks/{id}/rename", new { name });
        await SendAsync(req);
    }

    // ============ 考勤 ============

    public async Task<List<AttendanceTask>> GetAttendanceAsync(string? machineUuid = null, string? taskId = null)
    {
        var path = "/api/mobile/attendance";
        var query = new List<string>();
        if (!string.IsNullOrEmpty(machineUuid)) query.Add($"machine_uuid={Uri.EscapeDataString(machineUuid)}");
        if (!string.IsNullOrEmpty(taskId)) query.Add($"task_id={Uri.EscapeDataString(taskId)}");
        if (query.Count > 0) path += "?" + string.Join("&", query);
        using var req = MakeRequest(HttpMethod.Get, path);
        var json = await SendAsync(req);
        return JsonSerializer.Deserialize<AttendanceResponse>(json?.GetRawText() ?? "{}", _jsonOptions)?.Tasks ?? new();
    }

    // ============ 签到历史 ============

    public async Task<HistoryResponse> GetHistoryAsync()
    {
        using var req = MakeRequest(HttpMethod.Get, "/api/mobile/students/history");
        var json = await SendAsync(req);
        return JsonSerializer.Deserialize<HistoryResponse>(json?.GetRawText() ?? "{}", _jsonOptions) ?? new HistoryResponse();
    }

    // ============ 用户管理 ============

    public async Task<List<RemoteUserItem>> GetUsersAsync()
    {
        using var req = MakeRequest(HttpMethod.Get, "/api/users");
        var json = await SendAsync(req);
        return JsonSerializer.Deserialize<List<RemoteUserItem>>(json?.GetRawText() ?? "[]", _jsonOptions) ?? new();
    }

    public async Task CreateUserAsync(string username, string password, string role, string displayName)
    {
        using var req = MakeRequest(HttpMethod.Post, "/api/users",
            new { username, password, role, display_name = displayName });
        await SendAsync(req);
    }

    public async Task UpdateUserAsync(int id, string? role = null, bool? isActive = null, string? displayName = null)
    {
        using var req = MakeRequest(HttpMethod.Put, $"/api/users/{id}",
            new { role, is_active = isActive, display_name = displayName });
        await SendAsync(req);
    }

    public async Task DeleteUserAsync(int id)
    {
        using var req = MakeRequest(HttpMethod.Delete, $"/api/users/{id}");
        await SendAsync(req);
    }

    public async Task ChangePasswordAsync(string oldPassword, string newPassword, int? userId = null)
    {
        using var req = MakeRequest(HttpMethod.Post, "/api/users/change-password",
            new { old_password = oldPassword, new_password = newPassword, user_id = userId });
        await SendAsync(req);
    }
}
