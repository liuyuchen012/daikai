using System.IO;
using System.Net.Http;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using CheckIn.Shared.Models;

namespace CheckIn.Client.Services;

/// <summary>
/// 远程服务器通信服务，负责客户端与 SignWave 集控平台的交互
/// 功能包括：客户端注册、RSA 签名认证、数据同步与加载、配置同步
/// </summary>
public class ServerService
{
    private readonly HttpClient _http;
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
        _http.DefaultRequestHeaders.Remove("X-Server-Password");
        _http.DefaultRequestHeaders.Add("X-Server-Password", password);
    }

    /// <summary>
    /// 构造函数：初始化 HttpClient 并加载或创建 RSA 密钥对
    /// </summary>
    public ServerService()
    {
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
        LoadOrCreateKeys();
    }

    /// <summary>
    /// 加载已有的 RSA 密钥对和 UUID，如果不存在则创建新的并持久化到磁盘
    /// </summary>
    private void LoadOrCreateKeys()
    {
        var keyFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "client_key.pem");
        var uuidFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "client_uuid.txt");

        if (File.Exists(keyFile) && File.Exists(uuidFile))
        {
            // 从磁盘加载已有的密钥和UUID
            _rsa = RSA.Create();
            var keyPem = File.ReadAllText(keyFile);
            _rsa.ImportFromPem(keyPem);
            _clientUuid = File.ReadAllText(uuidFile).Trim();
        }
        else
        {
            // 生成新的 RSA 2048 密钥对和 UUID
            _rsa = RSA.Create(2048);
            var privKey = _rsa.ExportRSAPrivateKey();
            var pem = PemEncoding.Write("RSA PRIVATE KEY", privKey);
            File.WriteAllText(keyFile, pem);
            _clientUuid = Guid.NewGuid().ToString();
            File.WriteAllText(uuidFile, _clientUuid);
        }
        _publicKeyPem = _rsa.ExportSubjectPublicKeyInfoPem();
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
            var body = new { public_key = _publicKeyPem, name = machineName, password = _password, task_id = _taskId };
            var res = await _http.PostAsJsonAsync($"{_baseUrl}/api/register", body);
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
    /// 检查当前客户端在服务器上的在线状态
    /// </summary>
    public async Task<bool> CheckStatusAsync()
    {
        try
        {
            var res = await _http.GetAsync($"{_baseUrl}/api/status");
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
            var body = new { uuid = _clientUuid, task_id = _taskId, signature = Sign(dataStr), data = dataStr, password = _password };
            await _http.PostAsJsonAsync($"{_baseUrl}/api/sync_data", body);
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
            var res = await _http.PostAsJsonAsync($"{_baseUrl}/api/load_data", body);
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
            await _http.PostAsJsonAsync($"{_baseUrl}/api/update_config", body);
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
            var res = await _http.PostAsJsonAsync($"{_baseUrl}/api/get_config", body);
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
            var res = await _http.PostAsJsonAsync($"{_baseUrl}/api/create_signin", body);
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
            var res = await _http.PostAsJsonAsync($"{_baseUrl}/api/signin_result", body);
            if (!res.IsSuccessStatusCode) return null;
            return await res.Content.ReadAsStringAsync();
        }
        catch { return null; }
    }

    /// <summary>
    /// 获取服务器基础 URL（用于构建签到链接）
    /// </summary>
    public string BaseUrl => _baseUrl;
}
