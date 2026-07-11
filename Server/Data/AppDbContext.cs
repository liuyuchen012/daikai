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
/// 应用程序数据库上下文，使用 SQLite 存储设备信息和打卡记录
/// </summary>
public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

    /// <summary>设备表</summary>
    public DbSet<MachineEntity> Machines => Set<MachineEntity>();
    /// <summary>打卡记录表</summary>
    public DbSet<AttendanceEntity> AttendanceRecords => Set<AttendanceEntity>();

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
    }
}
