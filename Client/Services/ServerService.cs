using System.IO;
using System.Net.Http;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using CheckIn.Shared.Models;

namespace CheckIn.Client.Services;

public class ServerService
{
    private readonly HttpClient _http;
    private string _baseUrl = "";
    private string _password = "";
    private RSA? _rsa;
    private string? _clientUuid;
    private string? _publicKeyPem;

    public string? ClientUuid => _clientUuid;

    public void Initialize(string ip, int port, string password)
    {
        _baseUrl = $"http://{ip}:{port}";
        _password = password;
        _http.DefaultRequestHeaders.Remove("X-Server-Password");
        _http.DefaultRequestHeaders.Add("X-Server-Password", password);
    }

    public ServerService()
    {
        _http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
        LoadOrCreateKeys();
    }

    private void LoadOrCreateKeys()
    {
        var keyFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "client_key.pem");
        var uuidFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "client_uuid.txt");

        if (File.Exists(keyFile) && File.Exists(uuidFile))
        {
            _rsa = RSA.Create();
            var keyPem = File.ReadAllText(keyFile);
            _rsa.ImportFromPem(keyPem);
            _clientUuid = File.ReadAllText(uuidFile).Trim();
        }
        else
        {
            _rsa = RSA.Create(2048);
            var privKey = _rsa.ExportRSAPrivateKey();
            var pem = PemEncoding.Write("RSA PRIVATE KEY", privKey);
            File.WriteAllText(keyFile, pem);
            _clientUuid = Guid.NewGuid().ToString();
            File.WriteAllText(uuidFile, _clientUuid);
        }
        _publicKeyPem = _rsa.ExportSubjectPublicKeyInfoPem();
    }

    private string Sign(string message)
    {
        if (_rsa == null) throw new InvalidOperationException("RSA not initialized");
        var sig = _rsa.SignData(Encoding.UTF8.GetBytes(message), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        return Convert.ToBase64String(sig);
    }

    public async Task<bool> RegisterAsync(string machineName)
    {
        try
        {
            var body = new { public_key = _publicKeyPem, name = machineName, password = _password };
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

    public async Task SyncDataAsync(Dictionary<string, StudentAttendance> data)
    {
        try
        {
            var dataStr = JsonSerializer.Serialize(data);
            var body = new { uuid = _clientUuid, signature = Sign(dataStr), data = dataStr, password = _password };
            await _http.PostAsJsonAsync($"{_baseUrl}/api/sync_data", body);
        }
        catch { }
    }

    public async Task<Dictionary<string, StudentAttendance>?> LoadDataAsync()
    {
        try
        {
            var challenge = DateTime.Now.Ticks.ToString();
            var body = new { uuid = _clientUuid, signature = Sign(challenge), challenge = challenge, password = _password };
            var res = await _http.PostAsJsonAsync($"{_baseUrl}/api/load_data", body);
            if (!res.IsSuccessStatusCode) return null;
            var json = await res.Content.ReadFromJsonAsync<JsonElement>();
            var dataEl = json.GetProperty("data");
            return JsonSerializer.Deserialize<Dictionary<string, StudentAttendance>>(dataEl.GetRawText());
        }
        catch { return null; }
    }

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
            return JsonSerializer.Deserialize<ClientConfig>(configEl.GetRawText());
        }
        catch { return null; }
    }
}
