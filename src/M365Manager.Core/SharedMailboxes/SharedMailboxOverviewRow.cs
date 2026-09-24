namespace M365Manager.Core.SharedMailboxes;

/// <summary>One row of the Shared Mailboxes overview grid - read from the SQL cache (dbo.SharedMailboxes), never live from M365.</summary>
public sealed class SharedMailboxOverviewRow
{
    public string ExchangeGuid { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Alias { get; init; } = "";
    public string PrimarySmtpAddress { get; init; } = "";
    public string Owners { get; init; } = "";
    public string EDGroupAddress { get; init; } = "";
    public string AUGroupAddress { get; init; } = "";
    public string REGroupAddress { get; init; } = "";
    public bool? RequireSenderAuthenticationEnabled { get; init; }
    public bool IsDeletedInM365 { get; init; }
    public DateTime? CreatedDateTime { get; init; }
    public DateTime LastImportedAtUtc { get; init; }
}
