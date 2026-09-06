using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using CallCenter.Models;

namespace CallCenter.Services;

/// <summary>与服务端通信的 HTTP 客户端</summary>
public class ApiClient
{
    private readonly HttpClient _http = new() { Timeout = System.TimeSpan.FromSeconds(5) };

    public string BaseUrl { get; set; } = "http://127.0.0.1:5260";

    private string Normalized => BaseUrl.EndsWith('/') ? BaseUrl.TrimEnd('/') : BaseUrl;

    /// <summary>获取设备列表</summary>
    public async Task<List<DeviceInfo>> GetDevicesAsync()
    {
        var res = await _http.GetFromJsonAsync<DeviceListResponse>($"{Normalized}/api/devices");
        return res?.Devices ?? new List<DeviceInfo>();
    }

    /// <summary>发送呼叫</summary>
    public async Task<bool> SendCallAsync(DeviceInfo? target, string type, string title, string message, string sender)
    {
        var res = await _http.PostAsJsonAsync($"{Normalized}/api/calls/send", new
        {
            targetUuid = target?.Uuid,
            type,
            title,
            message,
            sender
        });
        return res.IsSuccessStatusCode;
    }

    /// <summary>测试服务端连通性</summary>
    public async Task<bool> TestConnectionAsync()
    {
        try
        {
            var res = await _http.GetAsync($"{Normalized}/api/devices");
            return res.IsSuccessStatusCode;
        }
        catch
        {
            return false;
        }
    }

    private class DeviceListResponse
    {
        public List<DeviceInfo>? Devices { get; set; }
    }
}
