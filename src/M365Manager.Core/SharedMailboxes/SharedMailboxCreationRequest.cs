namespace M365Manager.Core.SharedMailboxes;

/// <summary>
/// Input for <see cref="ISharedMailboxService.CreateSharedMailboxAsync"/> - mirrors the fields the
/// legacy "New Shared Mailbox" WinForms dialog (ShrMbxNewForm.ps1) asked for.
///
/// The .ED group is always created (with <see cref="OwnerIdentities"/> as its members when
/// <see cref="EditorMembers"/> is empty); .AU and .RE are only created when members are named for
/// them.
/// </summary>
public sealed class SharedMailboxCreationRequest
{
    /// <summary>ServiceNow (or equivalent) task number authorizing the creation. Required.</summary>
    public required string TaskNumber { get; init; }

    /// <summary>Location/site code used as the first segment of the enforced name.</summary>
    public required string Location { get; init; }

    /// <summary>Base mailbox name, before the "&lt;LOCATION&gt; " prefix is applied.</summary>
    public required string Name { get; init; }

    /// <summary>Verified tenant domain the mailbox address and access groups are created under.</summary>
    public required string Domain { get; init; }

    /// <summary>
    /// Comma-separated owner identities (SamAccountName, UPN or e-mail). Set as -ManagedBy on all
    /// three access groups - i.e. who is authorized to manage each group's membership.
    /// </summary>
    public required string OwnerIdentities { get; init; }

    /// <summary>
    /// Comma-separated identities to add as .ED (Editor) group members: FullAccess + SendAs on the
    /// mailbox. When empty, <see cref="OwnerIdentities"/> is used instead - the group itself is
    /// always created.
    /// </summary>
    public string EditorMembers { get; init; } = "";

    /// <summary>Comma-separated identities to add as .AU (Author) group members: send-on-behalf, no FullAccess. Empty means the group isn't created at all.</summary>
    public string AuthorMembers { get; init; } = "";

    /// <summary>Comma-separated identities to add as .RE (Reader) group members: read-only folder access. Empty means the group isn't created at all.</summary>
    public string ReaderMembers { get; init; } = "";

    /// <summary>If false, the mailbox only accepts mail from authenticated (internal) senders.</summary>
    public bool AllowExternalSenders { get; init; }

    /// <summary>
    /// Overrides the display name the convention would produce. Empty means "use the computed one".
    /// The legacy dialog pre-filled the computed name into an editable field for exactly this
    /// reason - some names simply can't be derived correctly from location + name.
    /// The access-group names follow whatever ends up here.
    /// </summary>
    public string DisplayNameOverride { get; init; } = "";

    /// <summary>
    /// Overrides the primary SMTP address the convention would produce. Empty means "use the
    /// computed one". The alias is always derived from this address and sanitized.
    /// </summary>
    public string AddressOverride { get; init; } = "";
}
