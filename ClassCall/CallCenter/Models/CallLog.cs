namespace CallCenter.Models;

/// <summary>呼叫记录（呼出端发送后展示）</summary>
public class CallLog
{
    public DateTime Time { get; set; } = DateTime.Now;
    public string Type { get; set; } = "notice";
    public string TypeText => Type switch
    {
        "urgent" => "紧急呼叫",
        "speech" => "仅朗读",
        _ => "普通通知"
    };
    public string Target { get; set; } = "全部设备";
    public string Title { get; set; } = "";
    public string Message { get; set; } = "";
    public string Sender { get; set; } = "";
    public bool Success { get; set; }

    public string TimeText => Time.ToString("HH:mm:ss");
    public string SuccessText => Success ? "成功" : "失败";
}
