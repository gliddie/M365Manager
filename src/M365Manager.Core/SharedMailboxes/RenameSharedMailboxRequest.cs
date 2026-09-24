namespace M365Manager.Core.SharedMailboxes;

/// <summary>
/// Input for <see cref="ISharedMailboxService.RenameAsync"/> - mirrors the legacy
/// RenameShrMbx.ps1 "Rename Shared Mailbox" dialog: the operator names the mailbox as it is now
/// plus the new location/name, the convention proposes a display name and address, and whatever
/// ends up in those (editable) fields is applied to the mailbox AND to every .ED/.AU/.RE access
/// group it has, so the group names keep matching the mailbox.
/// </summary>
public sealed class RenameSharedMailboxRequest
{
    /// <summary>ServiceNow (or equivalent) task number authorizing the rename. Required.</summary>
    public required string TaskNumber { get; init; }

    /// <summary>The mailbox as it is now: display name, alias, UPN, or any of its SMTP addresses.</summary>
    public required string MailboxIdentity { get; init; }

    /// <summary>New location/site code - first segment of the enforced name.</summary>
    public required string Location { get; init; }

    /// <summary>New base mailbox name, before the "&lt;LOCATION&gt; " prefix is applied.</summary>
    public required string Name { get; init; }

    /// <summary>
    /// Domain for the new address. Empty keeps the mailbox's current one - a rename is about the
    /// name, and moving a mailbox to another domain is a separate decision the operator has to
    /// make deliberately.
    /// </summary>
    public string Domain { get; init; } = "";

    /// <summary>Overrides the display name the convention would produce. Empty means "use the computed one".</summary>
    public string DisplayNameOverride { get; init; } = "";

    /// <summary>
    /// Overrides the primary SMTP address the convention would produce. Empty means "use the
    /// computed one". The alias is always re-derived from whichever address wins, and sanitized.
    /// </summary>
    public string AddressOverride { get; init; } = "";
}
