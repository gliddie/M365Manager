namespace M365Manager.Core.SharedMailboxes;

/// <summary>
/// Orchestrates creating and managing shared mailboxes the same way the legacy ShrMbxNew.ps1 /
/// ShrMbxChgOwner.ps1 / RenameShrMbx.ps1 / ShrMbxRemove.ps1 / RecoverShrMbx.ps1 WinForms scripts
/// did: enforced naming, three access-level
/// security groups (.ED/.AU/.RE) whose ManagedBy is the mailbox's owner(s) and whose membership
/// grants the actual Editor/Author/Reader mailbox permissions, MailTip/notes maintenance, an owner
/// confirmation e-mail (fully automated via Graph - unlike the legacy .oft template, which only
/// opened a draft for the operator to send manually), and a full step-by-step audit trail in
/// dbo.LogEntries.
/// </summary>
public interface ISharedMailboxService
{
    Task<SharedMailboxCreationResult> CreateSharedMailboxAsync(SharedMailboxCreationRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    Task<ChangeSharedMailboxOwnerResult> ChangeOwnerAsync(ChangeSharedMailboxOwnerRequest request, CancellationToken ct = default);

    /// <summary>
    /// Reads back what a rename would touch - the mailbox's current name/address/alias and the
    /// .ED/.AU/.RE access groups it actually has - without changing anything. The rename form shows
    /// this before the operator confirms, the way the legacy script's second dialog did.
    /// Throws when the mailbox cannot be resolved; the message is meant for the operator.
    /// </summary>
    Task<SharedMailboxRenamePreview> PreviewRenameAsync(string mailboxIdentity, CancellationToken ct = default);

    /// <summary>
    /// Renames the mailbox and every .ED/.AU/.RE access group it has, keeping the previous
    /// addresses as secondary aliases so mail sent to them keeps arriving.
    /// </summary>
    Task<RenameSharedMailboxResult> RenameAsync(RenameSharedMailboxRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>
    /// Reads back what a removal would take with it - owners, forwarding, access groups and their
    /// members, direct folder permissions - without changing anything. Throws when the mailbox
    /// cannot be resolved or is not a shared mailbox; the message is meant for the operator.
    /// </summary>
    Task<SharedMailboxRemovalPreview> PreviewRemovalAsync(string mailboxIdentity, CancellationToken ct = default);

    /// <summary>
    /// Writes a full snapshot to the audit log, then deletes the access groups (when asked) and
    /// the mailbox, and notifies the owners. Refuses anything that is not a shared mailbox.
    /// </summary>
    Task<RemoveSharedMailboxResult> RemoveAsync(RemoveSharedMailboxRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>Soft-deleted shared mailboxes, newest first - the ones still recoverable.</summary>
    Task<IReadOnlyList<DeletedSharedMailbox>> GetDeletedSharedMailboxesAsync(CancellationToken ct = default);

    /// <summary>
    /// The owners and members recorded when this app removed the mailbox, or null when it was
    /// removed some other way (or before the app wrote this data).
    /// </summary>
    Task<SharedMailboxRestoreData?> GetRestoreDataAsync(Guid exchangeGuid, CancellationToken ct = default);

    /// <summary>
    /// Restores a soft-deleted shared mailbox and recreates its access groups - deleted groups do
    /// not come back with it - then notifies the owners.
    /// </summary>
    Task<RecoverSharedMailboxResult> RecoverAsync(RecoverSharedMailboxRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>Verified, mail-enabled domains of the signed-in tenant, for the address-domain picker.</summary>
    Task<IReadOnlyList<string>> GetAvailableDomainsAsync(CancellationToken ct = default);

    /// <summary>Reads the Shared Mailboxes overview grid from the SQL cache (dbo.SharedMailboxes) - never live from M365.</summary>
    Task<IReadOnlyList<SharedMailboxOverviewRow>> GetSharedMailboxesOverviewAsync(CancellationToken ct = default);
}
