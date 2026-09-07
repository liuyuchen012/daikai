namespace CallPlugin.Models;

/// <summary>
/// 呼叫消息（来自服务端 /api/calls/pull）
/// Type: urgent 紧急呼叫 / notice 普通通知 / speech 仅朗读（不弹窗）
/// </summary>
public class CallMessage
{
    public int Id { get; set; }
    public string Type { get; set; } = "notice";
    public string Title { get; set; } = "";
    public string Message { get; set; } = "";
    public string Sender { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.Now;

    /// <summary>朗读文本（标题 + 内容拼合）</summary>
    public string SpeechText
    {
        get
        {
            var title = Title.Trim();
            var message = Message.Trim();
            if (title.Length == 0) return message;
            if (message.Length == 0) return title;
            return title.EndsWith("。") || title.EndsWith("！") || title.EndsWith("？")
                ? $"{title}{message}"
                : $"{title}。{message}";
        }
    }
}
