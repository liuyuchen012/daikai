using ClassIsland.Core.Abstractions.Services.SpeechService;

namespace CallPlugin.Services;

/// <summary>
/// ClassIsland 朗读（TTS）封装。
/// 通过 ClassIsland 主程序的 ISpeechService 将文本加入 TTS 队列朗读。
/// </summary>
public static class SpeechSpeaker
{
    /// <summary>
    /// 朗读文本。
    /// 服务容器通过 AppBase.Current.Services 反射获取（与主程序运行的程序集保持一致），
    /// 再强类型转换为 ISpeechService 调用 EnqueueSpeechQueue。
    /// </summary>
    public static void Speak(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return;
        try
        {
            var svc = ResolveSpeechService();
            svc?.EnqueueSpeechQueue(text);
        }
        catch
        {
            // 朗读失败不影响呼叫展示
        }
    }

    private static ISpeechService? ResolveSpeechService()
    {
        var app = ClassIsland.Core.AppBase.Current;
        if (app == null) return null;

        var services = app.GetType().GetProperty("Services")?.GetValue(app);
        if (services == null) return null;

        var getService = services.GetType().GetMethods()
            .FirstOrDefault(m => m.Name == "GetService" && m.GetParameters().Length == 1);
        if (getService == null) return null;

        return getService.Invoke(services, new object[] { typeof(ISpeechService) }) as ISpeechService;
    }
}
