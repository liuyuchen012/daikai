namespace CheckIn.Client.Models;

/// <summary>
/// 应用程序全局配置，持久化到 config.json 文件
/// 包含学校信息、按钮布局、服务器连接和管理员密码等设置
/// </summary>
public class AppConfig
{
    /// <summary>学校名称</summary>
    public string School { get; set; } = "";
    /// <summary>年级（年份）</summary>
    public string Nj { get; set; } = "";
    /// <summary>班级编号</summary>
    public string ClassId { get; set; } = "";
    /// <summary>课程名称（如"数学"）</summary>
    public string Km { get; set; } = "";
    /// <summary>学生按钮网格行数，默认 6 行</summary>
    public int ButtonRows { get; set; } = 6;
    /// <summary>学生按钮网格列数，默认 6 列</summary>
    public int ButtonCols { get; set; } = 6;
    /// <summary>是否启用在线模式（连接远程服务器同步）</summary>
    public bool OnlineMode { get; set; } = true;
    /// <summary>远程服务器 IP 地址</summary>
    public string ServerIp { get; set; } = "";
    /// <summary>远程服务器端口，默认 5250</summary>
    public int ServerPort { get; set; } = 5250;
    /// <summary>远程服务器连接密码</summary>
    public string ServerPassword { get; set; } = "";
    /// <summary>管理员密码的 SHA256 哈希值（空表示无密码）</summary>
    public string AdminPasswordHash { get; set; } = "";
    /// <summary>当前应用版本号（编译时常量，不参与序列化）</summary>
    public const string Version = "v2.8.0";
}
