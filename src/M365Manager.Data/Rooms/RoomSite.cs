namespace M365Manager.Data.Rooms;

/// <summary>
/// One site/location entry used when creating room and equipment mailboxes: maps the three-letter
/// site code that prefixes every room name to its calendar time zone and the regional RRS admin
/// group that gets FullAccess on rooms at that site.
/// Replaces the legacy "RoomTimeZones.csv" (E:\O365AdminShared\Data) that RoomResourceNewForm.ps1
/// read - unknown sites there were silently appended to the CSV; here they're maintained in
/// Settings.
/// </summary>
public sealed class RoomSite
{
    public int Id { get; set; }

    /// <summary>Three-letter site/office code, e.g. "NBK". Matched case-insensitively against the first word of the room name.</summary>
    public string SiteCode { get; set; } = "";

    /// <summary>Windows time zone id applied to the room mailbox, e.g. "W. Europe Standard Time".</summary>
    public string TimeZone { get; set; } = "";

    /// <summary>Region the site belongs to (legacy pick list: AP, CA, EU, LA, US).</summary>
    public string Region { get; set; } = "";

    /// <summary>Regional RRS admin group granted FullAccess on rooms at this site, e.g. "MBX.EU.RRS.Admins".</summary>
    public string RegionalAdminGroup { get; set; } = "";
}
