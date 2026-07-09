using Microsoft.EntityFrameworkCore;

namespace M365Manager.Data.Logging;

/// <summary>
/// EF Core-backed logging service writing to dbo.LogEntries on SQL Server.
/// A fresh DbContext is created per operation from the current connection string,
/// so changing the SQL settings at runtime takes effect immediately.
/// </summary>
public sealed class SqlLogService : ILogService
{
    private readonly IConnectionStringProvider _connectionStrings;

    public SqlLogService(IConnectionStringProvider connectionStrings)
    {
        _connectionStrings = connectionStrings;
    }

    private M365ManagerDbContext CreateContext()
    {
        var cs = _connectionStrings.GetConnectionString()
            ?? throw new InvalidOperationException("SQL Server is not configured. Open Settings and enter the connection details.");

        var options = new DbContextOptionsBuilder<M365ManagerDbContext>()
            .UseSqlServer(cs)
            .Options;

        return new M365ManagerDbContext(options);
    }

    public async Task WriteAsync(LogEntry entry, CancellationToken ct = default)
    {
        if (entry.TimestampUtc == default)
            entry.TimestampUtc = DateTime.UtcNow;
        entry.MachineName ??= Environment.MachineName;
        entry.WindowsUser ??= Environment.UserName;

        await using var ctx = CreateContext();
        ctx.LogEntries.Add(entry);
        await ctx.SaveChangesAsync(ct);
    }

    public async Task<IReadOnlyList<LogEntry>> GetRecentAsync(int count = 200, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        return await ctx.LogEntries
            .AsNoTracking()
            .OrderByDescending(x => x.Id)
            .Take(count)
            .ToListAsync(ct);
    }

    public async Task<bool> TestConnectionAsync(CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        return await ctx.Database.CanConnectAsync(ct);
    }
}
