namespace M365Manager.Data.Naming;

/// <summary>
/// CRUD access to the editable acronym-translation list used to build team names
/// (dbo.TeamNameAcronyms). Backs both the naming logic and the Settings management screen.
/// </summary>
public interface ITeamNamingRepository
{
    Task<IReadOnlyList<TeamNameAcronym>> GetAllAsync(CancellationToken ct = default);
    Task AddAsync(TeamNameAcronym acronym, CancellationToken ct = default);
    Task UpdateAsync(TeamNameAcronym acronym, CancellationToken ct = default);
    Task DeleteAsync(int id, CancellationToken ct = default);
}
