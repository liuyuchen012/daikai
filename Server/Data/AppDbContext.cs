using Microsoft.EntityFrameworkCore;

namespace CheckIn.Server.Data;

public class MachineEntity
{
    public string Uuid { get; set; } = Guid.NewGuid().ToString();
    public string Name { get; set; } = string.Empty;
    public string PublicKey { get; set; } = string.Empty;
    public string? LastSeen { get; set; }
    public string Config { get; set; } = "{}";
}

public class AttendanceEntity
{
    public int Id { get; set; }
    public string MachineUuid { get; set; } = string.Empty;
    public string Data { get; set; } = "{}";
    public string UpdatedAt { get; set; } = DateTime.Now.ToString("O");
}

public class AppDbContext : DbContext
{
    public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

    public DbSet<MachineEntity> Machines => Set<MachineEntity>();
    public DbSet<AttendanceEntity> AttendanceRecords => Set<AttendanceEntity>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<MachineEntity>(e =>
        {
            e.HasKey(m => m.Uuid);
            e.Property(m => m.Uuid).HasMaxLength(64);
        });

        modelBuilder.Entity<AttendanceEntity>(e =>
        {
            e.HasKey(a => a.Id);
            e.Property(a => a.MachineUuid).HasMaxLength(64);
            e.HasIndex(a => a.MachineUuid);
        });
    }
}
