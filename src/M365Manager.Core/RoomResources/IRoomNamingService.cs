namespace M365Manager.Core.RoomResources;

/// <summary>Every name/address the room creation flow derives from the user's input.</summary>
public sealed record RoomNames(
    string DisplayName,
    string LocalPart,
    string Address,
    string Alias,
    string RoomListName,
    string DelegateGroup,
    string UsersGroup,
    string DelegateGroupAddress,
    string UsersGroupAddress,
    string Office,
    string SiteCode);

/// <summary>
/// Builds room/equipment display names, addresses and the associated room-list, delegate and
/// users group names, matching RoomResourceNewForm.ps1::Build-MbxName and RoomEquipNew.ps1's
/// Office-string logic.
/// </summary>
public interface IRoomNamingService
{
    /// <summary>
    /// Applies the naming convention. <paramref name="rawName"/> is the "&lt;SITE&gt; &lt;Name&gt;"
    /// string the operator types; the first word is the site code. Known acronyms/terms
    /// (dbo.TeamNameAcronyms - the legacy scripts shared one KnownAcronyms.csv across all tools)
    /// are translated in the name part.
    /// </summary>
    Task<RoomNames> BuildAsync(
        string rawName,
        ResourceKind kind,
        RoomAccessModel accessModel,
        string building,
        string floor,
        string capacity,
        string mailboxDomain,
        string groupDomain,
        CancellationToken ct = default);
}
