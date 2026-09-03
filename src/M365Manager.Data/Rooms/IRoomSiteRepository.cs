namespace M365Manager.Data.Rooms;

/// <summary>
/// CRUD access to the editable site/time-zone list used when creating rooms (dbo.RoomSites).
/// Backs both the room creation lookup and the Settings management screen.
/// </summary>
public interface IRoomSiteRepository
{
    Task<IReadOnlyList<RoomSite>> GetAllAsync(CancellationToken ct = default);

    /// <summary>Case-insensitive lookup by site code; null when the site isn't maintained yet.</summary>
    Task<RoomSite?> FindByCodeAsync(string siteCode, CancellationToken ct = default);

    Task AddAsync(RoomSite site, CancellationToken ct = default);
    Task UpdateAsync(RoomSite site, CancellationToken ct = default);
    Task DeleteAsync(int id, CancellationToken ct = default);

    /// <summary>
    /// Inserts or updates by site code in one pass - used by the RoomTimeZones.csv import, so
    /// re-importing a corrected file doesn't produce duplicates.
    /// Returns how many rows were added and how many were updated.
    /// </summary>
    Task<(int Added, int Updated)> UpsertManyAsync(IEnumerable<RoomSite> sites, CancellationToken ct = default);
}
