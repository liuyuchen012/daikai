using CallPlugin.Services;
using CallPlugin.Views;
using ClassIsland.Core.Abstractions;
using ClassIsland.Core.Attributes;
using ClassIsland.Core.Extensions.Registry;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace CallPlugin;

/// <summary>
/// 班级呼叫系统 - 被控端插件入口
/// 功能：接收呼出端（控制端）通过服务端转发的班级呼叫，
/// 置顶弹窗提醒，并通过 ClassIsland 朗读功能朗读呼叫内容
/// </summary>
[PluginEntrance]
public class Plugin : PluginBase
{
    private CallReceiver? _receiver;

    public override void Initialize(HostBuilderContext context, IServiceCollection services)
    {
        // 注册插件设置页（ClassIsland 设置窗口 → 插件设置）
        PluginSettingsPage.ConfigFolder = PluginConfigFolder;
        services.AddSettingsPage<PluginSettingsPage>();

        _receiver = new CallReceiver(PluginConfigFolder);
        // 设置页保存后重启轮询以应用新配置
        PluginSettingsPage.OnSettingsSaved += () =>
        {
            _receiver.Stop();
            _receiver = new CallReceiver(PluginConfigFolder);
            _receiver.Start();
        };

        // 主程序启动完成后开始轮询
        var app = ClassIsland.Core.AppBase.Current;
        if (app != null)
            app.AppStarted += (_, _) => _receiver.Start();
        else
            _receiver.Start();
    }
}
