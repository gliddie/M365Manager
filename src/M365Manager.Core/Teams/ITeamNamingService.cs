namespace M365Manager.Core.Teams;

/// <summary>
/// Builds the enforced team display name (GRP.&lt;Location&gt;.&lt;Name&gt;) and its mail alias,
/// matching the legacy Check-GroupName / alias-building logic from NewUnifiedGrp.ps1.
/// </summary>
public interface ITeamNamingService
{
    /// <summary>
    /// Applies the naming convention: "GRP." + location + "." + name, then translates any
    /// known acronyms/terms (dbo.TeamNameAcronyms) found in the result.
    /// </summary>
    Task<string> BuildDisplayNameAsync(string location, string name, CancellationToken ct = default);

    /// <summary>
    /// Derives a mail-nickname-safe alias from a display name: strips spaces and the
    /// characters " /,$#_-\." and truncates to 64 characters (Exchange's alias limit).
    /// </summary>
    string BuildAlias(string displayName);
}
