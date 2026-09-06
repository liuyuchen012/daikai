using AgoraIn.ClassIslandPlugin.Models;
using Avalonia.Threading;
using ClassIsland.Core.Abstractions.Services.NotificationProviders;
using ClassIsland.Core.Attributes;
using ClassIsland.Core.Models.Notification;
using System.Text;

namespace AgoraIn.ClassIslandPlugin.Services.NotificationProviders;

/// <summary>
/// AgoraIn 呼叫提醒提供方（ClassIsland 标准提醒模式）
///
/// 继承 NotificationProviderBase：呼叫到达时调用基类 ShowNotification，
/// 弹出 ClassIsland 全屏遮罩提醒（Mask）→ 详情正文（Overlay）。
///
/// 内容使用官方工厂方法 CreateTwoIconsMask / CreateSimpleTextContent 构建，
/// 由类岛内置模板（TwoIconsMaskTemplate / SimpleTextOverlayTemplate）渲染，
/// 因此自动套用官方颜色逻辑（主题适配的 DynamicResource）与字号
/// （MainWindowEmphasizedFontSize 等），避免自绘控件在深色/浅色主题下不美观、
/// 以及硬编码超大字号导致字体显示不全的问题。
///
/// 注册方式：插件初始化必须使用 services.AddNotificationProvider&lt;CallNotificationProvider&gt;()
/// —— 基类构造会从 NotificationProviderRegistryService 按 [NotificationProviderInfo]
/// 查询注册信息，查不到会抛异常导致实例不创建（v2.4.0 不显示通知的根因）。
///
/// ShowNotification 必须由容器创建的唯一实例调用，因此通过静态 Instance 持有。
/// </summary>
[NotificationProviderInfo("B2E7A3C4-5D86-4F1A-9C0B-2A1F3D4E5C6D", "AgoraIn 呼叫提醒", "\ue138", "接收集控平台教师端发送的呼叫（待下课时段通知 / 上课应急通知 / 下课传唤），以 ClassIsland 提醒模式展示")]
public class CallNotificationProvider : NotificationProviderBase
{
    /// <summary>容器创建的唯一实例（托管服务单例），供轮询器调用</summary>
    public static CallNotificationProvider? Instance { get; private set; }

    /// <summary>课程服务（DI 注入，用于判断上课/下课状态；未注入时为 null）</summary>
    public static ClassIsland.Core.Abstractions.Services.ILessonsService? Lessons { get; private set; }

    public CallNotificationProvider(ClassIsland.Core.Abstractions.Services.ILessonsService lessons)
    {
        // 基类构造已完成：注册表查询 Info、Name/ProviderGuid/IconElement 初始化、
        // RegisterNotificationProvider(this) 注册到提醒主机
        Instance = this;
        Lessons = lessons;
    }

    /// <summary>
    /// 轮询器入口：任意线程可调用，内部切回 UI 线程后走 ShowNotification
    /// </summary>
    public static void Show(CallMessage call)
    {
        var provider = Instance;
        if (provider == null) return;

        Dispatcher.UIThread.Post(() =>
        {
            try
            {
                provider.ShowCall(call);
            }
            catch
            {
                // 提醒主机暂时不可用（如正在关闭）时静默降级
            }
        });
    }

    /// <summary>
    /// 按呼叫类型生成提醒（遮罩 + 正文），内容使用官方模板数据，
    /// 由类岛内置模板渲染，自动适配官方颜色与字号。
    /// </summary>
    // lucide 图标字形（与 ClassIsland 内置 Lucide 字体一致）
    private const string G_Siren = "\ue2ef";          // siren 警报
    private const string G_TriangleAlert = "\ue193";  // triangle-alert 三角警告
    private const string G_Megaphone = "\ue235";      // megaphone 喇叭
    private const string G_Bell = "\ue05d";           // bell 铃铛
    private const string G_BellRing = "\ue224";       // bell-ring 响铃

    private void ShowCall(CallMessage call)
    {
        var (label, leftIcon, rightIcon) = call.Type switch
        {
            "emergency" => ("上课应急通知", G_Siren, G_TriangleAlert),
            "summon"    => ("下课传唤", G_Megaphone, G_BellRing),
            _           => ("待下课时段通知", G_Bell, G_BellRing),
        };

        var maskText = $"{label}\n{call.Title}";
        if (!string.IsNullOrWhiteSpace(call.Message))
            maskText += $"\n{call.Message}";

        var overlayText = BuildOverlayText(call, label);

        var req = new NotificationRequest
        {
            // 官方遮罩：左右 lucide 图标 + 主题适配文字，图标/文字由官方模板渲染
            MaskContent = NotificationContent.CreateTwoIconsMask(
                maskText,
                leftIcon: $"lucide({leftIcon})",
                rightIcon: $"lucide({rightIcon})",
                hasRightIcon: true,
                factory: c => c.Duration = TimeSpan.FromSeconds(8)),

            // 官方正文：名单较长时用滚动文本，避免文字挤压/显示不全
            OverlayContent = call.Type == "summon" && !string.IsNullOrWhiteSpace(call.StudentNames)
                ? NotificationContent.CreateRollingTextContent(
                    overlayText,
                    duration: TimeSpan.FromSeconds(12),
                    repeatCount: 2,
                    factory: c => c.Duration = TimeSpan.FromSeconds(12))
                : NotificationContent.CreateSimpleTextContent(
                    overlayText,
                    factory: c => c.Duration = TimeSpan.FromSeconds(12)),
        };
        ShowNotification(req);

        // v2.5.0：朗读统一由 CallWindow.SpeakOnce 承担（三遍节奏，仅第一遍有声），
        // 避免与标准提醒重复发声。
    }

    /// <summary>正文文本：标题 + 内容 + 传唤名单（仅 summon）+ 发送者</summary>
    private static string BuildOverlayText(CallMessage call, string label)
    {
        var sb = new StringBuilder();
        sb.Append(label).Append('\n');
        sb.Append(call.Title);
        if (!string.IsNullOrWhiteSpace(call.Message))
            sb.Append('\n').Append(call.Message);
        if (call.Type == "summon" && !string.IsNullOrWhiteSpace(call.StudentNames))
        {
            var names = string.Join("、", call.StudentNames.Split(
                new[] { '\r', '\n', ',', '，' },
                StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
            sb.Append("\n\n请以下同学到办公室：").Append(names);
        }
        if (!string.IsNullOrWhiteSpace(call.Sender))
            sb.Append("\n\n发送者：").Append(call.Sender);
        return sb.ToString();
    }

    /// <summary>拼出 TTS 朗读文本</summary>
    private static string BuildSpokenText(CallMessage call, string label)
    {
        var sb = new StringBuilder();
        sb.Append(label).Append('，');
        sb.Append(call.Title);
        if (!string.IsNullOrWhiteSpace(call.Message))
            sb.Append('。').Append(call.Message);
        if (call.Type == "summon" && !string.IsNullOrWhiteSpace(call.StudentNames))
            sb.Append('。').Append("请以下同学到办公室：").Append(call.StudentNames.Replace('\n', '、'));
        return sb.ToString();
    }
}
