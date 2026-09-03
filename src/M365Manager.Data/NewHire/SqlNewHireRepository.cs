using Microsoft.EntityFrameworkCore;

namespace M365Manager.Data.NewHire;

/// <summary>
/// EF Core-backed access to the legacy telephony database. A fresh DbContext per operation, same
/// pattern as <see cref="M365Manager.Data.Rooms.SqlRoomSiteRepository"/>, so a changed connection
/// string takes effect immediately.
///
/// Every lookup here goes through LINQ rather than the concatenated SQL the legacy tool used - the
/// old queries interpolated the SamAccountName and office name straight into the statement.
/// </summary>
public sealed class SqlNewHireRepository : INewHireRepository
{
    private readonly INewHireConnectionStringProvider _connectionStrings;

    public SqlNewHireRepository(INewHireConnectionStringProvider connectionStrings)
    {
        _connectionStrings = connectionStrings;
    }

    public bool IsConfigured => _connectionStrings.GetConnectionString() is not null;

    private UcHelperDbContext CreateContext()
    {
        var cs = _connectionStrings.GetConnectionString()
            ?? throw new InvalidOperationException(
                "The telephony database is not configured. Open Settings > New Hire and enter its name.");

        var options = new DbContextOptionsBuilder<UcHelperDbContext>()
            .UseSqlServer(cs)
            .Options;

        return new UcHelperDbContext(options);
    }

    public async Task<bool> TestConnectionAsync(CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        if (!await ctx.Database.CanConnectAsync(ct))
            return false;

        // CanConnect only proves the catalog exists. Touching a table proves it is the right one.
        await ctx.Locations.AsNoTracking().Take(1).ToListAsync(ct);
        return true;
    }

    public async Task<IReadOnlyList<UcLocation>> GetLocationsAsync(CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        return await ctx.Locations
            .AsNoTracking()
            .Where(x => x.LocationCode != null && x.LocationCode != "")
            .OrderBy(x => x.Name)
            .ToListAsync(ct);
    }

    public async Task<UcLocation?> FindLocationByNameAsync(string officeName, CancellationToken ct = default)
    {
        var name = (officeName ?? "").Trim();
        if (name.Length == 0)
            return null;

        await using var ctx = CreateContext();
        return await ctx.Locations
            .AsNoTracking()
            .FirstOrDefaultAsync(x => x.Name == name, ct);
    }

    public async Task<UcEmployeeState> GetEmployeeStateAsync(string samAccountName, CancellationToken ct = default)
    {
        var sam = (samAccountName ?? "").Trim();
        if (sam.Length == 0)
            return new UcEmployeeState();

        await using var ctx = CreateContext();

        var endpoint = await ctx.Endpoints
            .AsNoTracking()
            .FirstOrDefaultAsync(x => x.SamAccountName == sam, ct);

        var license = await ctx.E3Licenses
            .AsNoTracking()
            .FirstOrDefaultAsync(x => x.SamAccountName == sam, ct);

        return new UcEmployeeState
        {
            Endpoint = endpoint,
            IsE3Licensed = license?.E3Licensed == "1",
            HasPhoneLicense = license?.HasPhoneLicense == "1",
        };
    }

    public async Task<IReadOnlyList<UcDidRange>> GetDidRangesAsync(string locationCode, CancellationToken ct = default)
    {
        var code = (locationCode ?? "").Trim();
        if (code.Length == 0)
            return Array.Empty<UcDidRange>();

        await using var ctx = CreateContext();
        return await ctx.DidRanges
            .AsNoTracking()
            .Where(x => x.LocationCode == code && x.Sdap == 1)
            .OrderBy(x => x.Utilization)
            .ThenBy(x => x.DidStart)
            .ToListAsync(ct);
    }

    public async Task<IReadOnlyList<FreeDid>> FindFreeDidsAsync(UcDidRange range, CancellationToken ct = default)
    {
        if (!long.TryParse(range.DidStart, out var start) || !long.TryParse(range.DidEnd, out var end))
            throw new InvalidOperationException(
                $"The range '{range.DidStart}-{range.DidEnd}' does not contain two numbers. Fix the row in the 'did' table.");

        if (end < start)
            throw new InvalidOperationException($"The range '{range.DidStart}-{range.DidEnd}' ends before it starts.");

        // A block wider than this is a data error, not a number range - enumerating it would hang
        // the UI and mean nothing. The legacy tool had no such guard.
        const long MaxRangeSize = 100_000;
        if (end - start >= MaxRangeSize)
            throw new InvalidOperationException(
                $"The range '{range.DidStart}-{range.DidEnd}' spans {end - start + 1} numbers, which is too wide to be a DID block.");

        // Everything in the block shares a prefix, so one LIKE narrows both lookups to the site
        // instead of pulling the whole endpoints table across. Same trick as the legacy DidFinder.
        var prefix = CommonPrefix(range.DidStart, range.DidEnd);

        await using var ctx = CreateContext();

        var assigned = await ctx.Endpoints
            .AsNoTracking()
            .Where(x => x.Did != null && x.Did.StartsWith(prefix))
            .Select(x => x.Did!)
            .ToListAsync(ct);

        var blocked = await ctx.BlockedDids
            .AsNoTracking()
            .Where(x => x.Did.StartsWith(prefix))
            .Select(x => x.Did)
            .ToListAsync(ct);

        var taken = new HashSet<string>(assigned.Concat(blocked), StringComparer.OrdinalIgnoreCase);

        var free = new List<FreeDid>();
        for (var number = start; number <= end; number++)
        {
            var text = number.ToString();
            if (!taken.Contains(text))
                free.Add(new FreeDid(text));
        }

        return free;
    }

    /// <summary>
    /// The leading digits both ends of a block share. The legacy tool always chopped the last five
    /// digits off the start number, which is wrong for any block that is not exactly five digits
    /// wide - too narrow a prefix silently misses assigned numbers and hands out a duplicate.
    /// </summary>
    private static string CommonPrefix(string start, string end)
    {
        var shared = 0;
        var max = Math.Min(start.Length, end.Length);
        while (shared < max && start[shared] == end[shared])
            shared++;

        return start[..shared];
    }

    public async Task<int> StartRunAsync(string samAccountName, string adminAccount, string ticket, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();

        var row = new UcNewHireLog
        {
            SamAccountName = samAccountName,
            AdminAccount = Truncate(adminAccount, 30),
            SnTicket = ticket,
            Date = DateTime.Now,
        };

        ctx.NewHireLogs.Add(row);
        await ctx.SaveChangesAsync(ct);
        return row.NewHireLogId;
    }

    public async Task CompleteRunAsync(int newHireLogId, Action<UcNewHireLog> apply, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();

        var row = await ctx.NewHireLogs.FirstOrDefaultAsync(x => x.NewHireLogId == newHireLogId, ct);
        if (row is null)
            return;

        apply(row);
        await ctx.SaveChangesAsync(ct);
    }

    public async Task AddEndpointAsync(UcEndpoint endpoint, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        ctx.Endpoints.Add(endpoint);
        await ctx.SaveChangesAsync(ct);
    }

    /// <summary>The legacy columns are narrow and not nullable; an over-long value would fail the insert.</summary>
    private static string Truncate(string? value, int max)
    {
        var text = (value ?? "").Trim();
        return text.Length <= max ? text : text[..max];
    }
}
