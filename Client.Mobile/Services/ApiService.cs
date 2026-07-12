using System.Net.Http.Json;
using System.Text;
using System.Text.Json;

namespace CheckIn.Client.Mobile.Services;

/// <summary>
/// 移动端 API 通信服务 - 使用 Bearer Token 认证
/// 替代旧的 RSA 密钥对认证方式
/// </summary>
public class ApiService
{
    private readonly HttpClient _httpClient;
    private string? _token;
    private string _baseUrl = "http://localhost:5250";

    public string? Token => _token;
    public string BaseUrl
    {
        get => _baseUrl;
        set => _baseUrl = value.TrimEnd('/');
    }

    public ApiService()
    {
        _httpClient = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(10)
        };
    }

    /// <summary>
    /// 设置认证令牌
    /// </summary>
    public void SetToken(string? token)
    {
        _token = token;
        _httpClient.DefaultRequestHeaders.Remove("Authorization");
        if (!string.IsNullOrEmpty(token))
            _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {token}");
    }

    /// <summary>
    /// 发送 POST JSON 请求
    /// </summary>
    public async Task<JsonElement> PostAsync(string endpoint, object? body = null)
    {
        var url = $"{_baseUrl}{endpoint}";
        StringContent? content = null;
        if (body != null)
        {
            var json = JsonSerializer.Serialize(body);
            content = new StringContent(json, Encoding.UTF8, "application/json");
        }
        var response = await _httpClient.PostAsync(url, content);
        var responseBody = await response.Content.ReadAsStringAsync();
        try
        {
            return JsonDocument.Parse(responseBody).RootElement;
        }
        catch
        {
            return JsonDocument.Parse("{}").RootElement;
        }
    }

    /// <summary>
    /// 发送 GET 请求
    /// </summary>
    public async Task<JsonElement> GetAsync(string endpoint)
    {
        var url = $"{_baseUrl}{endpoint}";
        var response = await _httpClient.GetAsync(url);
        var responseBody = await response.Content.ReadAsStringAsync();
        try
        {
            return JsonDocument.Parse(responseBody).RootElement;
        }
        catch
        {
            return JsonDocument.Parse("{}").RootElement;
        }
    }

    /// <summary>
    /// 发送 DELETE 请求
    /// </summary>
    public async Task<JsonElement> DeleteAsync(string endpoint)
    {
        var url = $"{_baseUrl}{endpoint}";
        var response = await _httpClient.DeleteAsync(url);
        var responseBody = await response.Content.ReadAsStringAsync();
        try
        {
            return JsonDocument.Parse(responseBody).RootElement;
        }
        catch
        {
            return JsonDocument.Parse("{}").RootElement;
        }
    }

    /// <summary>
    /// 从 JSON 响应中获取错误消息
    /// </summary>
    public static string? GetError(JsonElement json)
    {
        if (json.TryGetProperty("error", out var err))
            return err.GetString();
        return null;
    }

    /// <summary>
    /// 检查 JSON 响应是否成功（无 error 字段）
    /// </summary>
    public static bool IsSuccess(JsonElement json) => GetError(json) == null;

    /// <summary>
    /// 从 JSON 中获取字符串字段
    /// </summary>
    public static string? GetString(JsonElement json, string key)
    {
        return json.TryGetProperty(key, out var val) ? val.GetString() : null;
    }
}
