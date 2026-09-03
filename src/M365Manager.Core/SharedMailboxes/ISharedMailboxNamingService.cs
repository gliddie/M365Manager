namespace M365Manager.Core.SharedMailboxes;

/// <summary>
/// Builds shared mailbox display names, addresses and access-group names the same way the
/// legacy ShrMbxNewForm.ps1 "Build-MbxName" function did.
/// </summary>
public interface ISharedMailboxNamingService
{
    /// <summary>
    /// Applies the naming convention: location is upper-cased (title-cased instead if it starts
    /// with "Global"), the name is title-cased with known acronyms/terms (dbo.TeamNameAcronyms -
    /// the legacy scripts shared this same list between Teams and Shared Mailboxes) translated.
    /// Returns the display name ("&lt;LOCATION&gt; &lt;Name&gt;") and the address-safe local part
    /// ("&lt;LOCATION&gt;.&lt;Name&gt;", spaces/slashes/dashes stripped) used for the mailbox address.
    /// </summary>
    Task<(string DisplayName, string LocalPart)> BuildNameAsync(string location, string name, CancellationToken ct = default);

    string BuildAddress(string localPart, string domain);

    /// <summary>
    /// Name and display name of an access group: "MBX." + the mailbox's display name (location,
    /// a space, then the name with its spaces intact) + "." + tier.
    /// Example: mailbox "MXC Cristian Testet etwas" -> "MBX.MXC Cristian Testet etwas.ED".
    /// tier is "ED", "AU" or "RE".
    /// </summary>
    string BuildAccessGroupName(string mailboxDisplayName, string tier);

    /// <summary>
    /// Alias of an access group: <see cref="BuildAccessGroupName"/> with the spaces removed, since
    /// Exchange aliases and SMTP addresses can't contain them.
    /// Example: "MBX.MXCCristianTestetetwas.ED".
    /// </summary>
    string BuildAccessGroupAlias(string mailboxDisplayName, string tier);

    string BuildAccessGroupAddress(string mailboxDisplayName, string tier, string domain);
}
