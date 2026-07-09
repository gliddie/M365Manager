using Microsoft.EntityFrameworkCore;

namespace M365Manager.Data.Logging;

public sealed class M365ManagerDbContext : DbContext
{
    public M365ManagerDbContext(DbContextOptions<M365ManagerDbContext> options)
        : base(options)
    {
    }

    public DbSet<LogEntry> LogEntries => Set<LogEntry>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        var e = modelBuilder.Entity<LogEntry>();
        e.ToTable("LogEntries");
        e.HasKey(x => x.Id);

        e.Property(x => x.TimestampUtc).HasColumnType("datetime2(3)");
        e.Property(x => x.CreatedAtUtc).HasColumnType("datetime2(3)")
            .HasDefaultValueSql("SYSUTCDATETIME()");

        e.Property(x => x.UserUpn).HasMaxLength(256);
        e.Property(x => x.WindowsUser).HasMaxLength(128);
        e.Property(x => x.MachineName).HasMaxLength(128);
        e.Property(x => x.Area).HasMaxLength(64);
        e.Property(x => x.Action).HasMaxLength(128);
        e.Property(x => x.TargetObject).HasMaxLength(256);
        e.Property(x => x.EventCode).HasMaxLength(16);
        e.Property(x => x.Severity).HasColumnType("tinyint");

        e.HasIndex(x => x.TimestampUtc);
        e.HasIndex(x => x.UserUpn);
        e.HasIndex(x => x.TargetObject);
        e.HasIndex(x => x.Severity);
    }
}
