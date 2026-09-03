using M365Manager.Core.Naming;
using M365Manager.Data.Naming;

namespace M365Manager.Core.RoomResources;

/// <summary>
/// See <see cref="IRoomNamingService"/>. Reuses <see cref="ITeamNamingRepository"/> as-is
/// (dbo.TeamNameAcronyms) rather than its own table - the legacy scripts sourced Teams, shared
/// mailboxes and rooms from the same KnownAcronyms.csv.
/// </summary>
public sealed class RoomNamingService : IRoomNamingService
{
    private readonly ITeamNamingRepository _repository;

    public RoomNamingService(ITeamNamingRepository repository)
    {
        _repository = repository;
    }

    public async Task<RoomNames> BuildAsync(
        string rawName,
        ResourceKind kind,
        RoomAccessModel accessModel,
        string building,
        string floor,
        string capacity,
        string mailboxDomain,
        string groupDomain,
        CancellationToken ct = default)
    {
        // Used exactly as typed - no ToTitleCase. See SharedMailboxNamingService for why the
        // legacy title-casing (which mangles "IT" into "It") is not reproduced.
        var raw = (rawName ?? "").Trim();
        var space = raw.IndexOf(' ');

        // No space at all: treat the whole input as the site code with an empty name, so the
        // preview degrades gracefully while the operator is still typing (legacy threw here).
        // The site code is normalized and is part of every address and group name.
        var siteCode = ExchangeNaming.ToAliasSafe((space < 0 ? raw : raw[..space]).ToUpperInvariant());
        var namePart = space < 0 ? "" : raw[(space + 1)..];

        // Legacy quirk, preserved deliberately: the ADDRESS uses the name as typed (spaces
        // stripped, acronyms NOT expanded) while the DISPLAY NAME uses the expanded form - so
        // addresses stay short and display names stay readable. See
        // RoomResourceNewForm.ps1::Build-MbxName, where $RoomName is captured before the acronym
        // loop rewrites $MbxName.
        // Reduced to characters Exchange accepts in an alias, so "NBK M&A Meeting" or an umlaut
        // can't produce an address New-Mailbox rejects. The display name below keeps the original.
        var addressName = ExchangeNaming.ToAliasSafe(namePart.Replace(" ", ""));

        var acronyms = await _repository.GetAllAsync(ct);
        var expandedName = AcronymSubstitution.Apply(namePart, acronyms);

        var displayName = $"{siteCode} {expandedName}".TrimEnd();

        // Address local part: [RES.]<SITE>[Bldg<B>][.FLR<F>].<Name>[.<Capacity>]
        var localPart = kind == ResourceKind.Equipment ? $"RES.{siteCode}" : siteCode;

        var bldg = (building ?? "").Trim();
        localPart += bldg.Length > 0 ? $"Bldg{bldg}." : ".";

        var flr = (floor ?? "").Trim();
        if (flr.Length > 0)
        {
            localPart += IsGroundFloor(flr) ? "FLRGrnd" : $"FLR{flr}";
        }

        localPart += $".{addressName}";

        var cap = (capacity ?? "").Trim();
        if (kind == ResourceKind.Room && cap.Length > 0)
            localPart += $".{cap}";

        // Building, floor and capacity are free text too, so sanitize the assembled local part
        // once at the end - this also collapses the dot runs an empty segment leaves behind.
        localPart = ExchangeNaming.ToAliasSafe(localPart);

        var isRestricted = accessModel == RoomAccessModel.Restricted;
        var roomListName = isRestricted
            ? $"{siteCode} Restricted Rooms"
            : $"{siteCode} Conference Rooms";

        // Site-wide delegate group for the general-use default, per-room otherwise.
        //
        // The two legacy scripts disagreed on the per-room form: RoomResourceNewForm.ps1 used
        // DispName.Substring(4, DispName.IndexOf(" ") + 1), which slices an arbitrary few
        // characters out of the middle of the name ("NBK Meeting Room" -> "Meet"), while the newer
        // RoomResourceRename.ps1 used the full name with spaces stripped. The rename script's form
        // is the sane one and is what a renamed room ends up with anyway, so it wins here -
        // otherwise creating and then renaming a room would produce two different group names for
        // the same room.
        // These are used as both Name and Alias on the security group, so they go through the same
        // sanitizer - which also tidies the doubled dot an empty room name would leave behind.
        var delegateGroup = ExchangeNaming.ToAliasSafe(accessModel == RoomAccessModel.GeneralUseSiteDelegates
            ? $"MBX.{siteCode}.RRS.OutOfPolicy.DE"
            : $"MBX.{siteCode}.RRS.OutOfPolicy.{addressName}.DE");

        var usersGroup = isRestricted
            ? ExchangeNaming.ToAliasSafe($"MBX.{siteCode}.RRS.OutOfPolicy.{addressName}.US")
            : "";

        return new RoomNames(
            DisplayName: displayName,
            LocalPart: localPart,
            Address: $"{localPart}@{(mailboxDomain ?? "").Trim()}",
            Alias: localPart,
            RoomListName: roomListName,
            DelegateGroup: delegateGroup,
            UsersGroup: usersGroup,
            DelegateGroupAddress: $"{delegateGroup}@{(groupDomain ?? "").Trim()}",
            UsersGroupAddress: usersGroup.Length > 0 ? $"{usersGroup}@{(groupDomain ?? "").Trim()}" : "",
            Office: BuildOffice(building, floor),
            SiteCode: siteCode);
    }

    /// <summary>The Set-User -Office value, e.g. "Building 3, Floor 2" (RoomEquipNew.ps1:5-34).</summary>
    public static string BuildOffice(string? building, string? floor)
    {
        var bldg = (building ?? "").Trim();
        var flr = (floor ?? "").Trim();

        var floorText = flr.Length == 0 ? "" : IsGroundFloor(flr) ? "Ground Floor" : $"Floor {flr}";

        if (bldg.Length == 0)
            return floorText;

        return floorText.Length == 0 ? $"Building {bldg}" : $"Building {bldg}, {floorText}";
    }

    /// <summary>Legacy accepted both "0" (RoomEquipNew.ps1) and "Ground" (Build-MbxName) for the ground floor.</summary>
    private static bool IsGroundFloor(string floor)
        => floor.Equals("0", StringComparison.Ordinal)
        || floor.Equals("Ground", StringComparison.OrdinalIgnoreCase);

}
