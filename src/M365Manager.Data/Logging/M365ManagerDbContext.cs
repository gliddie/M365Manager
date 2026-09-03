using M365Manager.Data.Naming;
using M365Manager.Data.Rooms;
using Microsoft.EntityFrameworkCore;

namespace M365Manager.Data.Logging;

public sealed class M365ManagerDbContext : DbContext
{
    public M365ManagerDbContext(DbContextOptions<M365ManagerDbContext> options)
        : base(options)
    {
    }

    public DbSet<LogEntry> LogEntries => Set<LogEntry>();
    public DbSet<TeamNameAcronym> TeamNameAcronyms => Set<TeamNameAcronym>();
    public DbSet<RoomSite> RoomSites => Set<RoomSite>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        var acronym = modelBuilder.Entity<TeamNameAcronym>();
        acronym.ToTable("TeamNameAcronyms");
        acronym.HasKey(x => x.Id);
        acronym.Property(x => x.Acronym).HasMaxLength(64).IsRequired();
        acronym.Property(x => x.Translation).HasMaxLength(64).IsRequired();

        var site = modelBuilder.Entity<RoomSite>();
        site.ToTable("RoomSites");
        site.HasKey(x => x.Id);
        site.Property(x => x.SiteCode).HasMaxLength(16).IsRequired();
        site.Property(x => x.TimeZone).HasMaxLength(128).IsRequired();
        site.Property(x => x.Region).HasMaxLength(16);
        site.Property(x => x.RegionalAdminGroup).HasMaxLength(256);
        site.HasIndex(x => x.SiteCode).IsUnique();

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
        e.Property(x => x.TaskNumber).HasMaxLength(64);
        e.Property(x => x.EventCode).HasMaxLength(16);
        e.Property(x => x.Severity).HasColumnType("tinyint");

        e.HasIndex(x => x.TimestampUtc);
        e.HasIndex(x => x.UserUpn);
        e.HasIndex(x => x.TargetObject);
        e.HasIndex(x => x.Severity);
    }
}
