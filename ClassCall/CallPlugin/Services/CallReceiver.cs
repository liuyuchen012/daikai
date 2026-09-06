using System.Net.Http.Json;
using System.Text.Json;
using CallPlugin.Models;
using CallPlugin.Views;

namespace CallPlugin.Services;

/// <summary>
/// 呼叫接收器（被控端核心）：
/// 1. 定时从服务端拉取本设备的待处理呼叫
/// 2. 收到后朗读呼叫内容（ClassIsland TTS），urgent/notice 类型同时置顶弹窗
/// 3. 显示/朗读完成后向服务端确认（ack），避免重复拉取
/// </summary>
public class CallReceiver
{
    private readonly string _configFolder;
    private readonly PluginSettings _settings;
    private Timer? _timer;
    private volatile bool _polling;

    public CallReceiver(string configFolder)
    {
        _configFolder = configFolder;
        _settings = PluginSettings.Load(configFolder);
    }

    /// <summary>停止轮询（配置变更后由插件重启）</summary>
    public void Stop()
    {
        _timer?.Dispose();
        _timer = null;
    }

    public void Start()
    {
        // 尚未配置（无 UUID/密码）：向服务端注册本机设备，自动换取凭证
        if (!_settings.IsConfigured)
        {
            if (!TryAutoRegister()) return;
        }

        var interval = TimeSpan.FromSeconds(Math.Max(3, _settings.PollIntervalSeconds));
        _timer = new Timer(async _ => await PollAsync(), null, TimeSpan.FromSeconds(2), interval);
    }

    /// <summary>
    /// 首次使用自动注册：向服务端注册设备（名称/房间），换取 UUID + 密码并保存。
    /// 未填服务器地址或服务端不可达时返回 false（等待用户在设置页配置）。
    /// </summary>
    private bool TryAutoRegister()
    {
        try
        {
            if (string.IsNullOrEmpty(_settings.ServerUrl)) return false;

            var baseUrl = _settings.ServerUrl.EndsWith('/') ? _settings.ServerUrl : _settings.ServerUrl + "/";
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };

            var name = string.IsNullOrWhiteSpace(_settings.DeviceName) ? Environment.MachineName : _settings.DeviceName.Trim();
            var res = http.PostAsJsonAsync(baseUrl + "api/devices/register", new { name, room = _settings.Room ?? "" }).Result;
            if (!res.IsSuccessStatusCode) return false;

            using var doc = JsonDocument.Parse(res.Content.ReadAsStringAsync().Result);
            _settings.DeviceUuid = doc.RootElement.GetProperty("uuid").GetString() ?? "";
            _settings.Password = doc.RootElement.GetProperty("password").GetString() ?? "";
            _settings.DeviceName = name;
            _settings.Save(_configFolder);
            return _settings.IsConfigured;
        }
        catch
        {
            return false;
        }
    }

    private async Task PollAsync()
    {
        if (_polling) return;          // 防止上一轮未结束时重入
        _polling = true;
        try
        {
            var baseUrl = _settings.ServerUrl.EndsWith('/') ? _settings.ServerUrl : _settings.ServerUrl + "/";
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(6) };

            var res = await http.PostAsJsonAsync(baseUrl + "api/calls/pull",
                new { uuid = _settings.DeviceUuid, password = _settings.Password });
            if (!res.IsSuccessStatusCode) return;

            using var doc = JsonDocument.Parse(await res.Content.ReadAsStringAsync());
            foreach (var item in doc.RootElement.GetProperty("calls").EnumerateArray())
            {
                var call = new CallMessage
                {
                    Id = item.GetProperty("id").GetInt32(),
                    Type = item.GetProperty("type").GetString() ?? "notice",
                    Title = item.GetProperty("title").GetString() ?? "",
                    Message = item.GetProperty("message").GetString() ?? "",
                    Sender = item.GetProperty("sender").GetString() ?? "",
                    CreatedAt = item.TryGetProperty("createdAt", out var ca) ? ca.GetDateTime() : DateTime.Now
                };
                await ShowAndAckAsync(call, http, baseUrl);
            }
        }
        catch
        {
            // 服务端暂时不可达，下轮重试
        }
        finally
        {
            _polling = false;
        }
    }

    /// <summary>展示（朗读/弹窗）并确认呼叫</summary>
    private async Task ShowAndAckAsync(CallMessage call, HttpClient http, string baseUrl)
    {
        // 朗读呼叫内容（切到 UI 线程调用 ClassIsland TTS，urgent/notice 朗读标题+内容，speech 仅朗读不弹窗）
        if (_settings.SpeechEnabled)
        {
            var text = call.SpeechText;
            Avalonia.Threading.Dispatcher.UIThread.Post(() => SpeechSpeaker.Speak(text));
        }

        // speech 类型仅朗读，不弹窗
        if (call.Type != "speech")
            CallWindow.Show(call);

        try
        {
            await http.PostAsJsonAsync(baseUrl + "api/calls/ack",
                new { id = call.Id, uuid = _settings.DeviceUuid, password = _settings.Password });
        }
        catch
        {
            // 确认失败下轮会重复拉取，可接受
        }
    }
}
