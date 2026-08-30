using AgoraIn.ClassIslandPlugin.Services;
using AgoraIn.ClassIslandPlugin.Views;
using ClassIsland.Core.Abstractions;
using ClassIsland.Core.Attributes;
using ClassIsland.Core.Extensions.Registry;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace AgoraIn.ClassIslandPlugin;

/// <summary>
/// AgoraIn 联动插件入口
/// 功能：接收集控平台教师端发送的呼叫（待下课时段通知 / 上课应急通知 / 下课传唤）
/// 并结合 ClassIsland 课表在下课时间前自动提醒
/// </summary>
[PluginEntrance]
public class Plugin : PluginBase
{
    private CallPoller? _poller;

    public override void Initialize(HostBuilderContext context, IServiceCollection services)
    {
        // 注册插件设置页（ClassIsland 设置窗口 → 插件设置）
        PluginSettingsPage.ConfigFolder = PluginConfigFolder;
        services.AddSettingsPage<PluginSettingsPage>();

        _poller = new CallPoller(PluginConfigFolder);
        // 设置页保存后重启轮询以应用新配置
        PluginSettingsPage.OnSettingsSaved += () =>
        {
            _poller.Stop();
            _poller = new CallPoller(PluginConfigFolder);
            _poller.Start();
        };

        // 直接启动轮询，不依赖 AppStarted：在部分 ClassIsland 版本上该事件在插件订阅前已触发，
        // 导致 Start() 永不执行、呼叫永远收不到（用户反馈"安装后一直不显示"的根因）。
        // AppStarted 若仍触发，Start 内部会先 Stop 再启动，不会出现双轮询。
        _poller.Start();

        var app = ClassIsland.Core.AppBase.Current;
        if (app != null)
            app.AppStarted += (_, _) => _poller.Start();
    }
}
