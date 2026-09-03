using Microsoft.EntityFrameworkCore;
using M365Manager.Data.Logging;

namespace M365Manager.Data.Naming;

/// <summary>
/// EF Core-backed CRUD for dbo.TeamNameAcronyms. A fresh DbContext is created per operation,
/// same pattern as <see cref="SqlLogService"/>, so a changed connection string takes effect
/// immediately.
/// </summary>
public sealed class SqlTeamNamingRepository : ITeamNamingRepository
{
    private readonly IConnectionStringProvider _connectionStrings;

    public SqlTeamNamingRepository(IConnectionStringProvider connectionStrings)
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

    public async Task<IReadOnlyList<TeamNameAcronym>> GetAllAsync(CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        return await ctx.TeamNameAcronyms
            .AsNoTracking()
            .OrderBy(x => x.Acronym)
            .ToListAsync(ct);
    }

    public async Task AddAsync(TeamNameAcronym acronym, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        ctx.TeamNameAcronyms.Add(acronym);
        await ctx.SaveChangesAsync(ct);
    }

    public async Task UpdateAsync(TeamNameAcronym acronym, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        ctx.TeamNameAcronyms.Update(acronym);
        await ctx.SaveChangesAsync(ct);
    }

    public async Task DeleteAsync(int id, CancellationToken ct = default)
    {
        await using var ctx = CreateContext();
        var entity = await ctx.TeamNameAcronyms.FindAsync([id], ct);
        if (entity is not null)
        {
            ctx.TeamNameAcronyms.Remove(entity);
            await ctx.SaveChangesAsync(ct);
        }
    }
}
