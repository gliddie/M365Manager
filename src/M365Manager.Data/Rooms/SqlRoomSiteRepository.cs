using M365Manager.Data.Logging;
using Microsoft.EntityFrameworkCore;

namespace M365Manager.Data.Rooms;

/// <summary>
/// EF Core-backed CRUD for dbo.RoomSites. A fresh DbContext is created per operation, same pattern
/// as <see cref="M365Manager.Data.Naming.SqlTeamNamingRepository"/>, so a changed connection string
/// takes effect immediately.
/// </summary>
public sealed class SqlRoomSiteRepository : IRoomSiteRepository
{
    private readonly IConnectionStringProvider _connectionStrings;

    public SqlRoomSiteRepository(IConnectionStringProvider connectionStrings)
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

    public async Task<IReadOnlyList<RoomSite>> GetAllAsync(CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        return await ctx.RoomSites
            .AsNoTracking()
            .OrderBy(x => x.SiteCode)
            .ToListAsync(ct);
    }

    public async Task<RoomSite?> FindByCodeAsync(string siteCode, CancellationToken ct = default)
    {
        var code = (siteCode ?? "").Trim();
        if (code.Length == 0)
            return null;

        await using var ctx = CreateContext();
        // EF translates this to SQL, where the column's collation decides case sensitivity - which
        // is case-insensitive by default on SQL Server, matching the legacy CSV lookup.
        return await ctx.RoomSites
            .AsNoTracking()
            .FirstOrDefaultAsync(x => x.SiteCode == code, ct);
    }

    public async Task AddAsync(RoomSite site, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        ctx.RoomSites.Add(site);
        await ctx.SaveChangesAsync(ct);
    }

    public async Task UpdateAsync(RoomSite site, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        ctx.RoomSites.Update(site);
        await ctx.SaveChangesAsync(ct);
    }

    public async Task DeleteAsync(int id, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        var entity = await ctx.RoomSites.FindAsync([id], ct);
        if (entity is not null)
        {
            ctx.RoomSites.Remove(entity);
            await ctx.SaveChangesAsync(ct);
        }
    }

    public async Task<(int Added, int Updated)> UpsertManyAsync(IEnumerable<RoomSite> sites, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();

        var existing = await ctx.RoomSites.ToDictionaryAsync(x => x.SiteCode, StringComparer.OrdinalIgnoreCase, ct);
        var added = 0;
        var updated = 0;

        foreach (var site in sites)
        {
            var code = (site.SiteCode ?? "").Trim();
            if (code.Length == 0)
                continue;

            if (existing.TryGetValue(code, out var row))
            {
                row.TimeZone = site.TimeZone;
                row.Region = site.Region;
                row.RegionalAdminGroup = site.RegionalAdminGroup;
                updated++;
            }
            else
            {
                var entity = new RoomSite
                {
                    SiteCode = code,
                    TimeZone = site.TimeZone,
                    Region = site.Region,
                    RegionalAdminGroup = site.RegionalAdminGroup,
                };
                ctx.RoomSites.Add(entity);
                existing[code] = entity; // a duplicate code later in the file updates this row
                added++;
            }
        }

        await ctx.SaveChangesAsync(ct);
        return (added, updated);
    }
}
