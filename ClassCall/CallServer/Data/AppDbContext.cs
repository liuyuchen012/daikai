using CallServer.Models;
using Microsoft.EntityFrameworkCore;

namespace CallServer.Data;

public class AppDbContext(DbContextOptions<AppDbContext> options) : DbContext(options)
{
    public DbSet<Device> Devices => Set<Device>();
    public DbSet<CallRecord> CallRecords => Set<CallRecord>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        modelBuilder.Entity<Device>().HasIndex(d => d.Uuid).IsUnique();
    }
}
