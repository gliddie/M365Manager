namespace M365Manager.Core.SharedMailboxes;

/// <summary>How the given owner identities are applied to the mailbox's access groups' ManagedBy list.</summary>
public enum OwnerChangeMode
{
    Add,
    Remove,
    Replace,
}

/// <summary>
/// Input for <see cref="ISharedMailboxService.ChangeOwnerAsync"/> - mirrors ShrMbxChgOwner.ps1's
/// Add/Remove/Replace owner dialog. Applies to whichever of the mailbox's .ED/.AU/.RE access
/// groups exist.
/// </summary>
public sealed class ChangeSharedMailboxOwnerRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>Primary SMTP address of the shared mailbox.</summary>
    public required string MailboxAddress { get; init; }

    public required OwnerChangeMode Mode { get; init; }

    /// <summary>Comma-separated owner identities (SamAccountName, UPN or e-mail) to add, remove, or replace the current owners with.</summary>
    public required string OwnerIdentities { get; init; }
}
