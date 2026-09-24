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

    public async Task<IReadOnlyList<LogEntry>> QueryAsync(LogQuery query, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        var q = ctx.LogEntries.AsNoTracking();

        if (query.FromUtc is { } from)
            q = q.Where(x => x.TimestampUtc >= from);
        if (query.ToUtc is { } to)
            q = q.Where(x => x.TimestampUtc <= to);

        if (!string.IsNullOrWhiteSpace(query.UserUpn))
            q = q.Where(x => x.UserUpn == query.UserUpn);
        if (!string.IsNullOrWhiteSpace(query.Area))
            q = q.Where(x => x.Area == query.Area);
        if (!string.IsNullOrWhiteSpace(query.Action))
            q = q.Where(x => x.Action == query.Action);

        if (!string.IsNullOrWhiteSpace(query.TaskNumber))
        {
            var task = query.TaskNumber.Trim();
            q = q.Where(x => x.TaskNumber != null && x.TaskNumber.Contains(task));
        }

        if (query.Severity is { } severity)
            q = q.Where(x => x.Severity == severity);

        if (!string.IsNullOrWhiteSpace(query.SearchText))
        {
            var term = query.SearchText.Trim();
            q = q.Where(x =>
                (x.Message != null && x.Message.Contains(term))
                || (x.TargetObject != null && x.TargetObject.Contains(term))
                || (x.Action != null && x.Action.Contains(term))
                || (x.EventCode != null && x.EventCode.Contains(term)));
        }

        return await q
            .OrderByDescending(x => x.Id)
            .Take(query.MaxResults <= 0 ? 1000 : query.MaxResults)
            .ToListAsync(ct);
    }

    public async Task<LogFilterOptions> GetFilterOptionsAsync(CancellationToken ct = default)
    {
        await using var ctx = CreateContext();

        var users = await ctx.LogEntries.AsNoTracking()
            .Where(x => x.UserUpn != null && x.UserUpn != "")
            .Select(x => x.UserUpn!).Distinct().OrderBy(x => x).ToListAsync(ct);

        var areas = await ctx.LogEntries.AsNoTracking()
            .Where(x => x.Area != null && x.Area != "")
            .Select(x => x.Area!).Distinct().OrderBy(x => x).ToListAsync(ct);

        var actions = await ctx.LogEntries.AsNoTracking()
            .Where(x => x.Action != null && x.Action != "")
            .Select(x => x.Action!).Distinct().OrderBy(x => x).ToListAsync(ct);

        return new LogFilterOptions { Users = users, Areas = areas, Actions = actions };
    }

    public async Task<bool> TestConnectionAsync(CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        return await ctx.Database.CanConnectAsync(ct);
    }
}
