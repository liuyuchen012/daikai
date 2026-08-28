using System.Text.Json.Serialization;

namespace CheckIn.Shared.Models;

/// <summary>
/// 客户端配置模型（共享给 Client 和 Server 使用）
/// 用于在客户端与服务器之间同步任务配置和名称变更
/// </summary>
public class ClientConfig
{
    /// <summary>学校/任务名称</summary>
    public string School { get; set; } = string.Empty;
    /// <summary>年级</summary>
    public string Nj { get; set; } = string.Empty;
    /// <summary>班级</summary>
    public string ClassId { get; set; } = string.Empty;
    /// <summary>课程名称</summary>
    public string Km { get; set; } = string.Empty;
    /// <summary>按钮网格行数</summary>
    public int Z { get; set; } = 6;
    /// <summary>按钮网格列数</summary>
    public int L { get; set; } = 6;
    /// <summary>设备名称（服务端可修改后推送）</summary>
    public string DeviceName { get; set; } = string.Empty;
    /// <summary>配置版本号（自增，客户端通过版本号检测变更）</summary>
    public int ConfigVersion { get; set; }
    /// <summary>待推送的任务配置列表（服务端创建签到任务时写入，客户端应用后清空）</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public List<PendingTaskConfig>? PendingTasks { get; set; }
}

/// <summary>
/// 待推送的任务配置：服务端向设备推送的签到任务信息
/// </summary>
public class PendingTaskConfig
{
    /// <summary>短链码</summary>
    public string ShortCode { get; set; } = string.Empty;
    /// <summary>任务 ID（signin_xxx 格式）</summary>
    public string TaskId { get; set; } = string.Empty;
    /// <summary>科目名称</summary>
    public string Subject { get; set; } = string.Empty;
    /// <summary>教室名称</summary>
    public string Classroom { get; set; } = string.Empty;
    /// <summary>任务显示名称</summary>
    public string TaskName { get; set; } = string.Empty;
    /// <summary>签到密码</summary>
    public string Password { get; set; } = string.Empty;
    /// <summary>学生名单</summary>
    public List<string> Students { get; set; } = new();
    /// <summary>推送时间</summary>
    public string CreatedAt { get; set; } = string.Empty;
}
