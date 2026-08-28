namespace AgoraIn.ClassIslandPlugin.Models;

/// <summary>
/// 集控平台下发的呼叫消息
/// 三种类型：
/// prenotice - 待下课时段通知（提醒学生即将下课）
/// emergency - 上课应急通知（立即紧急播报）
/// summon    - 下课传唤（下课后叫学生）
/// </summary>
public class CallMessage
{
    /// <summary>服务端呼叫 ID</summary>
    public int Id { get; set; }

    /// <summary>呼叫类型：prenotice / emergency / summon</summary>
    public string Type { get; set; } = "prenotice";

    /// <summary>标题</summary>
    public string Title { get; set; } = "";

    /// <summary>内容</summary>
    public string Message { get; set; } = "";

    /// <summary>提前通知分钟数（仅 prenotice；0 = 到下课时间提醒）</summary>
    public int MinutesBefore { get; set; }

    /// <summary>传唤名单（仅 summon，换行或逗号分隔）</summary>
    public string StudentNames { get; set; } = "";

    /// <summary>发送者（教师）</summary>
    public string Sender { get; set; } = "";
}
