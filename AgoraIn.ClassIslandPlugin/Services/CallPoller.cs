using System.Net.Http.Json;
using System.Text.Json;
using AgoraIn.ClassIslandPlugin.Models;
using AgoraIn.ClassIslandPlugin.Services.NotificationProviders;

namespace AgoraIn.ClassIslandPlugin.Services;

/// <summary>
/// 呼叫轮询器：
/// 1. 定时从集控平台拉取本设备的待处理呼叫
/// 2. emergency / summon 立即显示；
///    prenotice（待下课通知）缓存起来，结合 ClassIsland 课表的下课剩余时间，在"距下课 ≤ 提前分钟数"时触发显示
/// 3. 显示后向平台确认（ack），避免重复拉取
/// </summary>
public class CallPoller
{
    private readonly string _configFolder;
    private readonly PluginSettings _settings;
    private readonly List<CallMessage> _pendingPrenotice = new();
    private Timer? _timer;

    public CallPoller(string configFolder)
    {
        _configFolder = configFolder;
        _settings = PluginSettings.Load(configFolder);
    }

    /// <summary>停止轮询（释放定时器），配置变更后由插件重启</summary>
    public void Stop()
    {
        _timer?.Dispose();
        _timer = null;
        lock (_pendingPrenotice)
            _pendingPrenotice.Clear();
    }

    public void Start()
    {
        // 幂等：重复调用（如 AppStarted 兜底触发）时先停掉旧定时器，避免双轮询
        _timer?.Dispose();
        _timer = null;

        if (string.IsNullOrEmpty(_settings.ServerUrl) || string.IsNullOrEmpty(_settings.Password))
        {
            // 平台地址/密码为空：尝试自动探测，仍为空则提示前往插件设置页配置
            PluginSettings.TryAutoDetect(_settings);
            _settings.Save(_configFolder);
            if (string.IsNullOrEmpty(_settings.ServerUrl) || string.IsNullOrEmpty(_settings.Password))
            {
                return;
            }
        }

        // 设备 UUID 为空：不再依赖本机 AgoraIn 客户端，自动向服务器登记为呼叫接收端
        if (string.IsNullOrEmpty(_settings.DeviceUuid))
        {
            var uuid = RegisterSelf(); // 同步尝试（失败时留空，由轮询重试）
            if (!string.IsNullOrEmpty(uuid))
            {
                _settings.DeviceUuid = uuid;
                _settings.Save(_configFolder);
            }
        }

        var interval = TimeSpan.FromSeconds(Math.Max(5, _settings.PollIntervalSeconds));
        _timer = new Timer(async _ => await PollAsync(), null, TimeSpan.FromSeconds(3), interval);
    }

    /// <summary>
    /// 呼叫接收端自登记：向服务器注册（无需 AgoraIn 客户端与 RSA 密钥），
    /// 教师端设备列表会用登记的 UUID 作为发送目标。广播呼叫（*）也总会送达。
    /// </summary>
    private string RegisterSelf()
    {
        try
        {
            var name = "ClassIsland-" + (System.Net.Dns.GetHostName() ?? "设备");
            using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(8) };
            var req = new HttpRequestMessage(HttpMethod.Post,
                new Uri((_settings.ServerUrl.EndsWith('/') ? _settings.ServerUrl : _settings.ServerUrl + "/") + "api/calls_register"))
            {
                Content = JsonContent.Create(new { name, password = _settings.Password, client_version = "ClassIslandPlugin-2.4.0" })
            };
            var res = http.Send(req);
            if (!res.IsSuccessStatusCode) return "";
            using var doc = System.Text.Json.JsonDocument.Parse(res.Content.ReadAsStringAsync().GetAwaiter().GetResult());
            return doc.RootElement.TryGetProperty("uuid", out var u) ? u.GetString() ?? "" : "";
        }
        catch { return ""; }
    }

    private async Task PollAsync()
    {
        try
        {
            using var http = new HttpClient();
            http.BaseAddress = new Uri(_settings.ServerUrl.EndsWith('/') ? _settings.ServerUrl : _settings.ServerUrl + "/");
            http.Timeout = TimeSpan.FromSeconds(6);

            var res = await http.PostAsJsonAsync("api/calls_pull", new
            {
                uuid = _settings.DeviceUuid,
                password = _settings.Password
            });
            if (!res.IsSuccessStatusCode)
            {
                CallStatusStore.SetStatus("无法连接集控平台");
                return;
            }

            CallStatusStore.SetStatus("已连接，正常监听");
            using var doc = JsonDocument.Parse(await res.Content.ReadAsStringAsync());
            foreach (var item in doc.RootElement.GetProperty("calls").EnumerateArray())
            {
                var call = new CallMessage
                {
                    Id = item.GetProperty("id").GetInt32(),
                    Type = item.GetProperty("type").GetString() ?? "prenotice",
                    Title = item.GetProperty("title").GetString() ?? "",
                    Message = item.GetProperty("message").GetString() ?? "",
                    MinutesBefore = item.TryGetProperty("minutes_before", out var mb) ? mb.GetInt32() : 0,
                    StudentNames = item.TryGetProperty("student_names", out var sn) ? sn.GetString() ?? "" : "",
                    Sender = item.TryGetProperty("sender", out var se) ? se.GetString() ?? "" : "",
                };

                if ((call.Type == "prenotice" && _settings.PrenoticeEnabled) || call.Type == "summon")
                {
                    // 待下课/下课传唤：按 ClassIsland 下课状态呼出——
                    // 上课时段（有课时且距下课 > 提前分钟数）暂缓；下课段/无课时立即呼出。
                    // 放假（无课表）时视为下课段立即呼出。
                    lock (_pendingPrenotice)
                        _pendingPrenotice.Add(call);
                }
                else
                {
                    await ShowAndAckAsync(call, http);
                }
            }

            // 检查待显示的待下课通知
            if (_settings.PrenoticeEnabled)
            {
                var left = GetBreakingTimeLeft();
                List<CallMessage>? toShow = null;
                lock (_pendingPrenotice)
                {
                    foreach (var call in _pendingPrenotice.ToList())
                    {
                        var trigger = left == null || left.Value.TotalMinutes <= Math.Max(0, call.MinutesBefore);
                        if (trigger)
                        {
                            _pendingPrenotice.Remove(call);
                            (toShow ??= new List<CallMessage>()).Add(call);
                        }
                    }
                }
                if (toShow != null)
                    foreach (var call in toShow)
                        await ShowAndAckAsync(call, http);
            }
        }
        catch
        {
            // 网络/平台暂时不可用，下轮重试
            CallStatusStore.SetStatus("连接出错，将自动重试");
        }
    }

    /// <summary>
    /// 显示呼叫并确认，防止重复拉取。
    /// 展示走 ClassIsland 标准提醒模式（提醒提供方 → 全屏遮罩 + 正文）。
    /// </summary>
    private async Task ShowAndAckAsync(CallMessage call, HttpClient http)
    {
        // ClassIsland 标准提醒（遮罩 + 正文）
        CallNotificationProvider.Show(call);
        // 更新主界面组件展示的最近呼叫
        CallStatusStore.SetLastCall(call);
        try
        {
            await http.PostAsJsonAsync("api/calls_ack", new
            {
                id = call.Id,
                uuid = _settings.DeviceUuid,
                password = _settings.Password
            });
        }
        catch { }
    }

    /// <summary>
    /// 获取 ClassIsland 课表的"距下课剩余时间"。
    /// 通过 ClassIsland.Core.AppBase.Current.Services 解析 IPublicLessonsService（反射，避免强依赖 IPC 程序集）
    /// 获取失败返回 null（此时待下课通知将立即显示）
    /// </summary>
    private static TimeSpan? GetBreakingTimeLeft()
    {
        try
        {
            var app = ClassIsland.Core.AppBase.Current;
            var servicesProp = app?.GetType().GetProperty("Services");
            var services = servicesProp?.GetValue(app);
            if (services == null) return null;

            var getService = services.GetType().GetMethods()
                .FirstOrDefault(m => m.Name == "GetService" && m.GetParameters().Length == 1);
            if (getService == null) return null;

            var ipcAsm = AppDomain.CurrentDomain.GetAssemblies()
                .FirstOrDefault(a => a.GetName().Name == "ClassIsland.Shared.IPC");
            var svcType = ipcAsm?.GetType("ClassIsland.Shared.IPC.Abstractions.Services.IPublicLessonsService");
            if (svcType == null) return null;

            var svc = getService.Invoke(services, new[] { svcType });
            if (svc == null) return null;

            var prop = svcType.GetProperty("OnBreakingTimeLeftTime");
            return prop?.GetValue(svc) is TimeSpan ts ? ts : null;
        }
        catch
        {
            return null;
        }
    }
}
