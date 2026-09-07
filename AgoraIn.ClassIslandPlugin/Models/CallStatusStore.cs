using Avalonia.Threading;

namespace AgoraIn.ClassIslandPlugin.Models;

/// <summary>
/// 呼叫状态单例：供主界面组件（CallStatusComponent）展示最近一次呼叫与连接状态。
/// 由 CallPoller 在轮询过程中更新（任意线程，内部切 UI 线程后触发事件）。
/// </summary>
public static class CallStatusStore
{
    /// <summary>最近一次呼叫（用于主界面组件展示）</summary>
    public static CallMessage? LastCall { get; private set; }

    /// <summary>最近一次轮询结果说明（如"已连接 / 未配置 / 网络错误"）</summary>
    public static string StatusText { get; private set; } = "等待首次轮询…";

    /// <summary>状态或最近呼叫变化时触发（在 UI 线程上）</summary>
    public static event Action? Changed;

    /// <summary>更新最近呼叫（任意线程调用）</summary>
    public static void SetLastCall(CallMessage? call)
    {
        Dispatcher.UIThread.Post(() =>
        {
            LastCall = call;
            Changed?.Invoke();
        });
    }

    /// <summary>更新状态说明（任意线程调用）</summary>
    public static void SetStatus(string text)
    {
        Dispatcher.UIThread.Post(() =>
        {
            StatusText = text;
            Changed?.Invoke();
        });
    }
}
