namespace M365Manager.Core.Teams;

/// <summary>One row of the Teams overview grid - read from the SQL cache (dbo.Teams + dbo.Groups), never live from M365.</summary>
public sealed class TeamOverviewRow
{
    public string GroupId { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Alias { get; init; } = "";
    public string PrimarySmtpAddress { get; init; } = "";
    public string Owners { get; init; } = "";
    public string Visibility { get; init; } = "";
    public bool IsArchived { get; init; }
    public bool? AllowToAddGuests { get; init; }
    public bool? HiddenFromAddressListsEnabled { get; init; }
    public bool? WelcomeMessageEnabled { get; init; }
    public string SharePointSiteUrl { get; init; } = "";
    public int MemberCount { get; init; }
    public bool IsDeletedInM365 { get; init; }
    public DateTime? CreatedDateTime { get; init; }
    public DateTime LastImportedAtUtc { get; init; }
}
