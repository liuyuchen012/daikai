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
    /// <summary>客户端版本号（客户端上报，如 v3.2.4）</summary>
    public string? ClientVersion { get; set; }
    /// <summary>客户端配置 JSON（学校、课程、待推送任务等）</summary>
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
    /// <summary>任务显示名称（可被管理员修改）</summary>
    public string TaskName { get; set; } = string.Empty;
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
/// 角色体系：admin（管理员）、teacher（普通教师）、student（学生）、parent（家长）
/// 兼容旧角色：operator 映射为 teacher，viewer 映射为 student
/// </summary>
public class UserEntity
{
    /// <summary>自增主键</summary>
    public int Id { get; set; }
    /// <summary>用户名（唯一）</summary>
    public string Username { get; set; } = string.Empty;
    /// <summary>密码哈希（加盐 SHA256）</summary>
    public string PasswordHash { get; set; } = string.Empty;
    /// <summary>角色：admin / teacher / student / parent（兼容 operator/viewer）</summary>
    public string Role { get; set; } = "student";
    /// <summary>显示名称</summary>
    public string DisplayName { get; set; } = string.Empty;
    /// <summary>创建时间（ISO 8601 格式）</summary>
    public string CreatedAt { get; set; } = DateTime.Now.ToString("O");
    /// <summary>是否启用</summary>
    public bool IsActive { get; set; } = true;
}

/// <summary>
/// 设备分配实体：管理员将设备分配给教师，可精细到任务级别
/// TaskId 为空表示分配该设备的所有任务
/// </summary>
public class DeviceAssignmentEntity
{
    /// <summary>自增主键</summary>
    public int Id { get; set; }
    /// <summary>被分配的教师用户 ID</summary>
    public int UserId { get; set; }
    /// <summary>设备 UUID</summary>
    public string MachineUuid { get; set; } = string.Empty;
    /// <summary>任务 ID（可选，为空则该设备所有任务可见）</summary>
    public string? TaskId { get; set; }
    /// <summary>分配者用户名</summary>
    public string AssignedBy { get; set; } = string.Empty;
    /// <summary>创建时间</summary>
    public string CreatedAt { get; set; } = DateTime.Now.ToString("O");
}

/// <summary>
/// 呼叫实体：教师通过集控平台向大屏设备发送的即时通知
/// 三种类型：prenotice（待下课时段通知）/ emergency（上课应急通知）/ summon（下课传唤）
/// </summary>
public class CallEntity
{
    /// <summary>自增主键</summary>
    public int Id { get; set; }
    /// <summary>呼叫类型：prenotice / emergency / summon</summary>
    public string Type { get; set; } = "prenotice";
    /// <summary>目标设备 UUID</summary>
    public string MachineUuid { get; set; } = string.Empty;
    /// <summary>标题</summary>
    public string Title { get; set; } = string.Empty;
    /// <summary>内容</summary>
    public string Message { get; set; } = string.Empty;
    /// <summary>提前通知分钟数（仅 prenotice 类型；0 表示当前下课即提醒）</summary>
    public int MinutesBefore { get; set; }
    /// <summary>传唤名单（仅 summon 类型，换行或逗号分隔）</summary>
    public string StudentNames { get; set; } = string.Empty;
    /// <summary>发送者（教师用户名）</summary>
    public string Sender { get; set; } = string.Empty;
    /// <summary>创建时间（ISO 8601）</summary>
    public string CreatedAt { get; set; } = DateTime.Now.ToString("O");
    /// <summary>状态：pending / acknowledged / expired</summary>
    public string Status { get; set; } = "pending";
    /// <summary>过期时间（ISO 8601），默认 2 小时后自动过期</summary>
    public string ExpiresAt { get; set; } = DateTime.Now.AddHours(2).ToString("O");
    /// <summary>重复遍数：设备需连续播报的总遍数（默认 1）。每次 ack 后若剩余遍数 &gt; 0 会自动克隆下一条 pending</summary>
    public int RepeatCount { get; set; } = 1;
}

/// <summary>
/// 系统日志实体：记录服务器关键操作（登录、发送呼叫、删除记录、配置变更等），
/// 供管理员在 Web 面板「系统日志」页面查看与删除。
/// </summary>
public class SystemLogEntity
{
    /// <summary>自增主键</summary>
    public int Id { get; set; }
    /// <summary>日志级别：info / warning / error</summary>
    public string Level { get; set; } = "info";
    /// <summary>操作类型（如 login / send_call / delete_call / delete_log / config 等）</summary>
    public string Category { get; set; } = string.Empty;
    /// <summary>操作者（用户名；系统操作为 system）</summary>
    public string Operator { get; set; } = "system";
    /// <summary>日志内容</summary>
    public string Message { get; set; } = string.Empty;
    /// <summary>创建时间（ISO 8601）</summary>
    public string CreatedAt { get; set; } = DateTime.Now.ToString("O");
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
    /// <summary>设备分配表</summary>
    public DbSet<DeviceAssignmentEntity> DeviceAssignments => Set<DeviceAssignmentEntity>();
    /// <summary>呼叫表</summary>
    public DbSet<CallEntity> Calls => Set<CallEntity>();
    /// <summary>系统日志表</summary>
    public DbSet<SystemLogEntity> SystemLogs => Set<SystemLogEntity>();

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
            e.Property(s => s.TaskName).HasMaxLength(128);
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

        // 设备分配实体配置
        modelBuilder.Entity<DeviceAssignmentEntity>(e =>
        {
            e.HasKey(d => d.Id);
            e.Property(d => d.MachineUuid).HasMaxLength(64);
            e.Property(d => d.TaskId).HasMaxLength(64);
            e.Property(d => d.AssignedBy).HasMaxLength(64);
            e.HasIndex(d => d.UserId);
            e.HasIndex(d => new { d.UserId, d.MachineUuid });
        });

        // 呼叫实体配置：设备+状态复合索引，优化设备拉取待处理呼叫
        modelBuilder.Entity<CallEntity>(e =>
        {
            e.HasKey(c => c.Id);
            e.Property(c => c.MachineUuid).HasMaxLength(64);
            e.Property(c => c.Type).HasMaxLength(16);
            e.Property(c => c.Status).HasMaxLength(16);
            e.HasIndex(c => new { c.MachineUuid, c.Status });
        });

        // 系统日志实体配置：按时间倒序查询（Id 倒序即可），级别与操作类型加索引
        modelBuilder.Entity<SystemLogEntity>(e =>
        {
            e.HasKey(l => l.Id);
            e.Property(l => l.Level).HasMaxLength(16);
            e.Property(l => l.Category).HasMaxLength(32);
            e.Property(l => l.Operator).HasMaxLength(64);
            e.HasIndex(l => l.Category);
            e.HasIndex(l => l.CreatedAt);
        });
    }
}
