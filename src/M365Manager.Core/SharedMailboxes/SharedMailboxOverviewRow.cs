namespace M365Manager.Core.SharedMailboxes;

/// <summary>One row of the Shared Mailboxes overview grid - read from the SQL cache (dbo.SharedMailboxes), never live from M365.</summary>
public sealed class SharedMailboxOverviewRow
{
    public string ExchangeGuid { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Alias { get; init; } = "";
    public string PrimarySmtpAddress { get; init; } = "";
    /// <summary>One entry per owner, "First Last &lt;upn&gt;" - rows the import has not rewritten yet may still hold a bare name or UPN.</summary>
    public IReadOnlyList<string> OwnerList { get; init; } = [];
    /// <summary>One line for the grid, the filter and the Excel export.</summary>
    public string Owners => string.Join(", ", OwnerList);
    /// <summary>One owner per line for the detail panel, which is too narrow for two side by side.</summary>
    public string OwnersMultiline => string.Join(Environment.NewLine, OwnerList);
    public string EDGroupAddress { get; init; } = "";
    public string AUGroupAddress { get; init; } = "";
    public string REGroupAddress { get; init; } = "";
    public bool? RequireSenderAuthenticationEnabled { get; init; }
    public bool IsDeletedInM365 { get; init; }
    public DateTime? CreatedDateTime { get; init; }
    public DateTime LastImportedAtUtc { get; init; }
}
