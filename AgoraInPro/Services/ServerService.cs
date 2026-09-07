using System.IO;
using System.Net.Http;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using CheckIn.Client.Models;
using CheckIn.Shared.Models;

namespace CheckIn.Client.Services;

/// <summary>
/// 远程服务器通信服务，负责客户端与 AgoraIn 集控平台的交互
/// 功能包括：客户端注册、RSA 签名认证、数据同步与加载、配置同步
/// </summary>
public class ServerService : IDisposable
{
    // M2：所有 ServerService 实例共享同一个静态 HttpClient，
    // 避免每个标签页都创建独立连接导致 Socket 端口耗尽。
    // 注意：不在此共享实例上设置 DefaultRequestHeaders（会互相污染），
    // 密码等认证头由 MakeRequest 为每次请求单独构造。
    private static readonly HttpClient _http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
    private string _baseUrl = "";
    private string _password = "";
    private RSA? _rsa;
    private string? _clientUuid;
    private string? _publicKeyPem;
    private string _taskId = "default";
    private static readonly JsonSerializerOptions _jsonOptions = new() { PropertyNameCaseInsensitive = true };

    /// <summary>客户端唯一标识符（UUID）</summary>
    public string? ClientUuid => _clientUuid;
    /// <summary>当前任务 ID，用于区分同一设备的不同打卡任务</summary>
    public string TaskId { get => _taskId; set => _taskId = value ?? "default"; }

    /// <summary>
    /// 初始化服务器连接参数（地址、端口、密码）
    /// </summary>
    public void Initialize(string ip, int port, string password)
    {
        _baseUrl = $"http://{ip}:{port}";
        _password = password;
    }

    /// <summary>
    /// 构造函数：加载或创建 RSA 密钥对（私钥以 DPAPI 加密存储，见 LoadOrCreateKeys）
    /// </summary>
    public ServerService()
    {
        LoadOrCreateKeys();
    }

    /// <summary>
    /// 加载已有的 RSA 密钥对和 UUID，如果不存在则创建新的并持久化到磁盘
    /// H2：私钥不再明文存储，改用 DPAPI（ProtectedData）按当前用户加密后落盘，
    /// 其他用户即使读取到文件也无法解密出私钥。
    /// 对旧版本遗留的明文 client_key.pem 会自动迁移为加密存储。
    /// </summary>
    private void LoadOrCreateKeys()
    {
        var keyFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "client_key.pem");
        var uuidFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "client_uuid.txt");

        // 使用文件锁防止并发实例创建重复密钥
        var lockFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "client_key.lock");
        using (var fs = new FileStream(lockFile, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
        {
            if (File.Exists(keyFile) && File.Exists(uuidFile))
            {
                // 从磁盘加载已有的密钥和UUID
                _rsa = RSA.Create();
                try
                {
                    var encrypted = File.ReadAllBytes(keyFile);
                    var keyPem = Encoding.UTF8.GetString(
                        ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser));
                    _rsa.ImportFromPem(keyPem);
                }
                catch (CryptographicException)
                {
                    // 兼容旧版本：文件为明文 PEM，读取后立即迁移为 DPAPI 加密存储
                    var keyPem = File.ReadAllText(keyFile);
                    _rsa.ImportFromPem(keyPem);
                    SavePrivateKeyEncrypted(keyFile, _rsa);
                }
                _clientUuid = File.ReadAllText(uuidFile).Trim();
            }
            else
            {
                // 生成新的 RSA 2048 密钥对和 UUID
                _rsa = RSA.Create(2048);
                SavePrivateKeyEncrypted(keyFile, _rsa);
                _clientUuid = Guid.NewGuid().ToString();
                File.WriteAllText(uuidFile, _clientUuid);
            }
        }
        _publicKeyPem = _rsa.ExportSubjectPublicKeyInfoPem();
    }

    /// <summary>将 RSA 私钥以 DPAPI（当前用户）加密后写入文件，绝不落明文</summary>
    private static void SavePrivateKeyEncrypted(string keyFile, RSA rsa)
    {
        var privKey = rsa.ExportRSAPrivateKey();
        var pem = new string(PemEncoding.Write("RSA PRIVATE KEY", privKey));
        var encrypted = ProtectedData.Protect(Encoding.UTF8.GetBytes(pem), null, DataProtectionScope.CurrentUser);
        File.WriteAllBytes(keyFile, encrypted);
    }

    /// <summary>
    /// 构造带认证头的 HTTP 请求。密码头按实例单独设置，避免共享 HttpClient 的默认头互相污染。
    /// </summary>
    private HttpRequestMessage MakeRequest(HttpMethod method, string url, object? body = null)
    {
        var req = new HttpRequestMessage(method, url);
        req.Headers.Add("X-Server-Password", _password);
        if (body != null)
            req.Content = JsonContent.Create(body);
        return req;
    }

    /// <summary>
    /// 使用 RSA 私钥对消息进行 SHA256 签名，返回 Base64 编码的签名
    /// </summary>
    private string Sign(string message)
    {
        if (_rsa == null) throw new InvalidOperationException("RSA not initialized");
        var sig = _rsa.SignData(Encoding.UTF8.GetBytes(message), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        return Convert.ToBase64String(sig);
    }

    /// <summary>
    /// 向服务器注册客户端，发送公钥、设备名称和任务 ID
    /// </summary>
    public async Task<bool> RegisterAsync(string machineName)
    {
        try
        {
            var body = new { public_key = _publicKeyPem, name = machineName, password = _password, task_id = _taskId, client_version = Models.AppConfig.Version };
            using var req = MakeRequest(HttpMethod.Post, $"{_baseUrl}/api/register", body);
            var res = await _http.SendAsync(req);
            if (!res.IsSuccessStatusCode) return false;
            var json = await res.Content.ReadFromJsonAsync<JsonElement>();
            var serverUuid = json.GetProperty("uuid").GetString()!;
            if (serverUuid != _clientUuid)
            {
                _clientUuid = serverUuid;
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "client_uuid.txt"), serverUuid);
            }
            return true;
        }
        catch { return false; }
    }

    /// <summary>
    /// 查询集控服务器告知的最新客户端版本（服务端转发 GitHub Release 资产）
    /// </summary>
    public async Task<(bool HasUpdate, string LatestVersion, string DownloadUrl)?> GetClientUpdateAsync()
    {
        try
        {
            var req = MakeRequest(HttpMethod.Post, $"{_baseUrl}/api/client_update", new { password = _password });
            var res = await _http.SendAsync(req);
            if (!res.IsSuccessStatusCode) return null;
            var json = await res.Content.ReadFromJsonAsync<JsonElement>();
            if (!json.TryGetProperty("has_update", out var h) || !h.GetBoolean()) return null;
            var latest = json.TryGetProperty("latest_version", out var lv) ? lv.GetString() ?? "" : "";
            var url = json.TryGetProperty("download_url", out var du) ? du.GetString() ?? "" : "";
            return (true, latest, url);
        }
        catch { return null; }
    }

    /// <summary>
    /// 检查当前客户端在服务器上的在线状态
    /// </summary>
    public async Task<bool> CheckStatusAsync()
    {
        try
        {
            using var req = MakeRequest(HttpMethod.Get, $"{_baseUrl}/api/status");
            var res = await _http.SendAsync(req);
            if (!res.IsSuccessStatusCode) return false;
            var machines = await res.Content.ReadFromJsonAsync<List<JsonElement>>();
            if (machines == null) return false;
            return machines.Any(m => m.GetProperty("uuid").GetString() == _clientUuid && m.GetProperty("online").GetBoolean());
        }
        catch { return false; }
    }

    /// <summary>
    /// 将本地打卡数据同步到服务器（含 RSA 签名验证）
    /// </summary>
    public async Task SyncDataAsync(Dictionary<string, StudentAttendance> data)
    {
        try
        {
            var dataStr = JsonSerializer.Serialize(data);
            var body = new { uuid = _clientUuid, task_id = _taskId, signature = Sign(dataStr), data = dataStr, password = _password, client_version = Models.AppConfig.Version };
            using var req = MakeRequest(HttpMethod.Post, $"{_baseUrl}/api/sync_data", body);
            await _http.SendAsync(req);
        }
        catch { }
    }

    /// <summary>
    /// 从服务器加载最新打卡数据（通过 challenge-签名机制验证身份）
    /// </summary>
    public async Task<Dictionary<string, StudentAttendance>?> LoadDataAsync()
    {
        try
        {
            var challenge = DateTime.Now.Ticks.ToString();
            var body = new { uuid = _clientUuid, task_id = _taskId, signature = Sign(challenge), challenge = challenge, password = _password };
            using var req = MakeRequest(HttpMethod.Post, $"{_baseUrl}/api/load_data", body);
            var res = await _http.SendAsync(req);
            if (!res.IsSuccessStatusCode) return null;
            var json = await res.Content.ReadFromJsonAsync<JsonElement>();
            var dataEl = json.GetProperty("data");
            return JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(dataEl.GetRawText(), _jsonOptions);
        }
        catch { return null; }
    }

    /// <summary>
    /// 将客户端配置（学校、课程、行列数）同步到服务器
    /// </summary>
    public async Task SyncConfigAsync(ClientConfig config)
    {
        try
        {
            var configStr = JsonSerializer.Serialize(config);
            var body = new { uuid = _clientUuid, signature = Sign(configStr), config, password = _password };
            using var req = MakeRequest(HttpMethod.Post, $"{_baseUrl}/api/update_config", body);
            await _http.SendAsync(req);
        }
        catch { }
    }

    /// <summary>
    /// 从服务器加载客户端配置
    /// </summary>
    public async Task<ClientConfig?> LoadConfigAsync()
    {
        try
        {
            var challenge = DateTime.Now.Ticks.ToString();
            var body = new { uuid = _clientUuid, signature = Sign(challenge), challenge, password = _password };
            using var req = MakeRequest(HttpMethod.Post, $"{_baseUrl}/api/get_config", body);
            var res = await _http.SendAsync(req);
            if (!res.IsSuccessStatusCode) return null;
            var json = await res.Content.ReadFromJsonAsync<JsonElement>();
            var configEl = json.GetProperty("config");
            return JsonSerializer.Deserialize<ClientConfig>(configEl.GetRawText(), _jsonOptions);
        }
        catch { return null; }
    }

    /// <summary>
    /// 创建远程签到任务：上传签到密码、教室、科目和学生名单，获取短链码
    /// </summary>
    /// <param name="signPassword">学生签到密码</param>
    /// <param name="classroom">教室名称</param>
    /// <param name="subject">科目名称</param>
    /// <param name="students">学生姓名列表</param>
    /// <returns>包含 short_code（短链码）和 task_id 的结果，失败返回 null</returns>
    public async Task<(string shortCode, string taskId)?> CreateSignInAsync(string signPassword, string classroom, string subject, List<string> students)
    {
        try
        {
            var body = new
            {
                uuid = _clientUuid,
                sign_password = signPassword,
                classroom,
                subject,
                students,
                password = _password
            };
            using var req = MakeRequest(HttpMethod.Post, $"{_baseUrl}/api/create_signin", body);
            var res = await _http.SendAsync(req);
            if (!res.IsSuccessStatusCode) return null;
            var json = await res.Content.ReadFromJsonAsync<JsonElement>();
            var shortCode = json.GetProperty("short_code").GetString() ?? "";
            var taskId = json.GetProperty("task_id").GetString() ?? "";
            return (shortCode, taskId);
        }
        catch { return null; }
    }

    /// <summary>
    /// 拉取签到任务的结果（学生签到记录），需 challenge-签名验证
    /// </summary>
    /// <returns>签到结果列表（JSON 字符串），失败返回 null</returns>
    public async Task<string?> GetSignInResultAsync()
    {
        try
        {
            var challenge = DateTime.Now.Ticks.ToString();
            var body = new { uuid = _clientUuid, signature = Sign(challenge), challenge, password = _password };
            using var req = MakeRequest(HttpMethod.Post, $"{_baseUrl}/api/signin_result", body);
            var res = await _http.SendAsync(req);
            if (!res.IsSuccessStatusCode) return null;
            return await res.Content.ReadAsStringAsync();
        }
        catch { return null; }
    }

    /// <summary>
    /// 获取服务器基础 URL（用于构建签到链接）
    /// </summary>
    public string BaseUrl => _baseUrl;

    /// <summary>
    /// 释放实例资源（H1）。
    /// HttpClient 为静态共享实例，不能在此释放；
    /// 该方法供 TaskTabViewModel.Dispose 明确标记连接已失效。
    /// </summary>
    public void Dispose() { }

    /// <summary>
    /// 客户端确认已应用推送的配置任务
    /// </summary>
    /// <param name="taskIds">已应用的任务 ID 列表</param>
    public async Task ConfigAppliedAsync(List<string> taskIds)
    {
        try
        {
            var body = new { uuid = _clientUuid, applied_tasks = taskIds, password = _password };
            using var req = MakeRequest(HttpMethod.Post, $"{_baseUrl}/api/config_applied", body);
            await _http.SendAsync(req);
        }
        catch { }
    }

    /// <summary>
    /// 查询集控平台自身是否有新版本（服务器后台定时检查 GitHub 的结果）
    /// </summary>
    /// <returns>元组 (hasUpdate, latestVersion, downloadUrl)，失败返回 null</returns>
    public async Task<(bool hasUpdate, string latestVersion, string downloadUrl)?> GetServerUpdateAsync()
    {
        try
        {
            using var req = MakeRequest(HttpMethod.Get, $"{_baseUrl}/api/server_update");
            var res = await _http.SendAsync(req);
            if (!res.IsSuccessStatusCode) return null;
            var json = await res.Content.ReadFromJsonAsync<JsonElement>();
            var hasUpdate = json.GetProperty("has_update").GetBoolean();
            var latestVersion = json.GetProperty("latest_version").GetString() ?? "";
            var downloadUrl = json.GetProperty("download_url").GetString() ?? "";
            return (hasUpdate, latestVersion, downloadUrl);
        }
        catch { return null; }
    }

    /// <summary>
    /// 拉取集控平台发给本设备的待处理呼叫（待下课通知 / 上课应急 / 下课传唤）
    /// </summary>
    public async Task<List<CallMessage>> PullCallsAsync()
    {
        try
        {
            using var req = MakeRequest(HttpMethod.Post, $"{_baseUrl}/api/calls_pull", new { uuid = _clientUuid, password = _password });
            var res = await _http.SendAsync(req);
            if (!res.IsSuccessStatusCode) return new();
            var json = await res.Content.ReadFromJsonAsync<JsonElement>();
            var calls = new List<CallMessage>();
            foreach (var item in json.GetProperty("calls").EnumerateArray())
                calls.Add(new CallMessage
                {
                    Id = item.GetProperty("id").GetInt32(),
                    Type = item.GetProperty("type").GetString() ?? "prenotice",
                    Title = item.GetProperty("title").GetString() ?? "",
                    Message = item.GetProperty("message").GetString() ?? "",
                    MinutesBefore = item.TryGetProperty("minutes_before", out var mb) ? mb.GetInt32() : 0,
                    StudentNames = item.TryGetProperty("student_names", out var sn) ? sn.GetString() ?? "" : "",
                    Sender = item.TryGetProperty("sender", out var se) ? se.GetString() ?? "" : "",
                });
            return calls;
        }
        catch { return new(); }
    }

    /// <summary>
    /// 确认已收到并显示呼叫，防止重复拉取
    /// </summary>
    public async Task AckCallAsync(int id)
    {
        try
        {
            using var req = MakeRequest(HttpMethod.Post, $"{_baseUrl}/api/calls_ack", new { id, uuid = _clientUuid, password = _password });
            var res = await _http.SendAsync(req);
            if (res.IsSuccessStatusCode)
            {
                var json = await res.Content.ReadFromJsonAsync<JsonElement>();
                if (json.TryGetProperty("status", out var st) && st.GetString() == "ok") { /* acked */ }
            }
        }
        catch { /* 确认失败下次轮询重试 */ }
    }
}
