namespace CheckIn.Shared.Models;

/// <summary>
/// 客户端配置模型（共享给 Client 和 Server 使用）
/// 用于在客户端与服务器之间同步任务配置
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
}
