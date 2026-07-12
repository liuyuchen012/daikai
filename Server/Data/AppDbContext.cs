using Microsoft.EntityFrameworkCore;

namespace CheckIn.Server.Data;

/// <summary>
/// 设备实体：存储注册到服务器的客户端设备信息
/// </summary>
public class MachineEntity
{
    /// <summary>设备唯一标识符</summary>
    public string Uuid { get; set; } = Guid.NewGuid().ToString();
    /// <summary>设备名称（如"三年（1）班"）</summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>客户端 RSA 公钥（PEM 格式），用于验证签名</summary>
    public string PublicKey { get; set; } = string.Empty;
    /// <summary>最后在线时间（ISO 8601 格式）</summary>
    public string? LastSeen { get; set; }
    /// <summary>客户端配置 JSON（学校、课程等）</summary>
    public string Config { get; set; } = "{}";
}

/// <summary>
/// 打卡记录实体：存储客户端提交的打卡数据
/// 按版本链方式追加记录，支持历史回溯
/// </summary>
public class AttendanceEntity
{
    /// <summary>自增主键</summary>
    public int Id { get; set; }
    /// <summary>所属设备 UUID</summary>
    public string MachineUuid { get; set; } = string.Empty;
    /// <summary>任务 ID，区分同一设备的不同打卡任务</summary>
    public string TaskId { get; set; } = "default";
    /// <summary>打卡数据 JSON（字典格式：学生名 -> StudentAttendance）</summary>
    public string Data { get; set; } = "{}";
    /// <summary>更新时间（ISO 8601 格式）</summary>
    public string UpdatedAt { get; set; } = DateTime.Now.ToString("O");
}

/// <summary>
/// 签到任务实体：存储教师创建的远程签到任务，学生通过短链页面签到
/// </summary>
public class SignInTaskEntity
{
    /// <summary>自增主键</summary>
    public int Id { get; set; }
    /// <summary>短链码（6-8位随机字符），作为签到页面路径标识</summary>
    public string ShortCode { get; set; } = string.Empty;
    /// <summary>创建该任务的设备 UUID</summary>
    public string MachineUuid { get; set; } = string.Empty;
    /// <summary>签到密码（学生签到需要输入）</summary>
    public string Password { get; set; } = string.Empty;
    /// <summary>教室名称</summary>
    public string Classroom { get; set; } = string.Empty;
    /// <summary>科目名称</summary>
    public string Subject { get; set; } = string.Empty;
    /// <summary>学生名单 JSON 数组（从 CSV 导入的学生姓名列表）</summary>
    public string StudentList { get; set; } = "[]";
    /// <summary>签到记录 JSON 数组（格式：[{name, time}]）</summary>
    public string SignInRecords { get; set; } = "[]";
    /// <summary>创建时间（ISO 8601 格式）</summary>
    public string CreatedAt { get; set; } = DateTime.Now.ToString("O");
    /// <summary>任务状态：active（进行中）/ closed（已关闭）</summary>
    public string Status { get; set; } = "active";
}

/// <summary>
/// 用户实体：存储 Web 管理面板的用户账号信息
/// </summary>
public class UserEntity
{
    /// <summary>自增主键</summary>
    public int Id { get; set; }
    /// <summary>用户名（唯一）</summary>
    public string Username { get; set; } = string.Empty;
    /// <summary>密码哈希（SHA256 十六进制字符串）</summary>
    public string PasswordHash { get; set; } = string.Empty;
    /// <summary>角色：admin（管理员）/ operator（操作员）/ viewer（查看者）</summary>
    public string Role { get; set; } = "viewer";
    /// <summary>显示名称</summary>
    public string DisplayName { get; set; } = string.Empty;
    /// <summary>创建时间（ISO 8601 格式）</summary>
    public string CreatedAt { get; set; } = DateTime.Now.ToString("O");
    /// <summary>是否启用</summary>
    public bool IsActive { get; set; } = true;
}

/// <summary>
/// 应用程序数据库上下文，使用 SQLite 存储设备信息、打卡记录和用户信息
/// </summary>
public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

    /// <summary>设备表</summary>
    public DbSet<MachineEntity> Machines => Set<MachineEntity>();
    /// <summary>打卡记录表</summary>
    public DbSet<AttendanceEntity> AttendanceRecords => Set<AttendanceEntity>();
    /// <summary>签到任务表</summary>
    public DbSet<SignInTaskEntity> SignInTasks => Set<SignInTaskEntity>();
    /// <summary>用户表</summary>
    public DbSet<UserEntity> Users => Set<UserEntity>();

    /// <summary>
    /// 配置实体映射：设置主键、字段长度限制和索引
    /// </summary>
    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        // 设备实体配置
        modelBuilder.Entity<MachineEntity>(e =>
        {
            e.HasKey(m => m.Uuid);
            e.Property(m => m.Uuid).HasMaxLength(64);
        });

        // 打卡记录实体配置：自增主键 + 复合索引（设备+任务）优化查询
        modelBuilder.Entity<AttendanceEntity>(e =>
        {
            e.HasKey(a => a.Id);
            e.Property(a => a.MachineUuid).HasMaxLength(64);
            e.Property(a => a.TaskId).HasMaxLength(64);
            e.HasIndex(a => a.MachineUuid);                       // 按设备查询索引
            e.HasIndex(a => new { a.MachineUuid, a.TaskId });     // 按设备+任务复合索引
        });

        // 签到任务实体配置
        modelBuilder.Entity<SignInTaskEntity>(e =>
        {
            e.HasKey(s => s.Id);
            e.Property(s => s.ShortCode).HasMaxLength(16);
            e.Property(s => s.MachineUuid).HasMaxLength(64);
            e.HasIndex(s => s.ShortCode).IsUnique();              // 短链码唯一索引
            e.HasIndex(s => s.MachineUuid);                       // 按设备查询索引
        });

        // 用户实体配置
        modelBuilder.Entity<UserEntity>(e =>
        {
            e.HasKey(u => u.Id);
            e.Property(u => u.Username).HasMaxLength(64);
            e.Property(u => u.PasswordHash).HasMaxLength(128);
            e.HasIndex(u => u.Username).IsUnique();               // 用户名唯一索引
        });
    }
}
