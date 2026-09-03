namespace M365Manager.Core.Groups;

/// <summary>
/// Administrative operations on distribution, security and dynamic groups, replacing the legacy
/// "Distribution List / Security and O365 Group Admin Menu" (DistributionSecurityGroupMenu.ps1 and
/// the ~10 scripts it dispatched to).
///
/// Deliberately a sibling of <see cref="M365Manager.Core.Exchange.IExchangeService"/> rather than an
/// extension of it: that service is also the IM365Connector and is consumed by other features,
/// while these are write operations with their own audit trail. Searching for groups and reading or
/// changing their membership stays there.
/// </summary>
public interface IGroupAdminService
{
    /// <summary>Resolves the derived names and checks what already exists, without writing anything.</summary>
    Task<GroupCreationPreview> PreviewAsync(GroupCreationRequest request, CancellationToken ct = default);

    /// <summary>Creates a distribution or mail-enabled security group with owners, members, aliases and authorized senders.</summary>
    Task<GroupOperationResult> CreateAsync(GroupCreationRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>Creates a dynamic distribution group from the given recipient conditions.</summary>
    Task<GroupOperationResult> CreateDynamicAsync(DynamicGroupCreationRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>Adds, removes or replaces the group's owners (ManagedBy, or UnifiedGroupLinks for M365 groups).</summary>
    Task<GroupOperationResult> ChangeOwnersAsync(GroupListChangeRequest request, CancellationToken ct = default);

    /// <summary>Adds, removes or replaces who may send to the group (AcceptMessagesOnlyFromSendersOrMembers).</summary>
    Task<GroupOperationResult> ChangeAuthorizedSendersAsync(GroupListChangeRequest request, CancellationToken ct = default);

    /// <summary>Adds or removes secondary SMTP addresses. Replace is not supported - it would drop the primary address.</summary>
    Task<GroupOperationResult> ChangeAliasesAsync(GroupListChangeRequest request, CancellationToken ct = default);

    /// <summary>Replaces the entire membership in one call (Update-DistributionGroupMember).</summary>
    Task<GroupOperationResult> ReplaceMembersAsync(GroupListChangeRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>Renames the group, keeping the previous address as a secondary alias.</summary>
    Task<GroupOperationResult> RenameAsync(GroupRenameRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>Captures a snapshot into the audit log, then deletes the group with whichever cmdlet fits its type.</summary>
    Task<GroupRemovalResult> RemoveAsync(GroupRemovalRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>
    /// Rebuilds the group's MailTip from its current owners, without changing anything else.
    /// Repairs tips left behind by the flattened-ManagedBy bug, which wrote raw directory values
    /// instead of names and was only ever corrected by making another owner change.
    /// </summary>
    Task<GroupOperationResult> RefreshMailTipAsync(string groupIdentity, string taskNumber, CancellationToken ct = default);

    /// <summary>Current owners, authorized senders and aliases of a group, for the manage panel.</summary>
    Task<GroupDetails?> GetDetailsAsync(string groupIdentity, CancellationToken ct = default);
}

/// <summary>The mutable settings the manage panel shows for the selected group.</summary>
public sealed class GroupDetails
{
    public string DisplayName { get; init; } = "";
    public string PrimarySmtpAddress { get; init; } = "";
    public string RecipientTypeDetails { get; init; } = "";
    public IReadOnlyList<string> Owners { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> AuthorizedSenders { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> Aliases { get; init; } = Array.Empty<string>();
    public bool RequireSenderAuthenticationEnabled { get; init; }
    public string Notes { get; init; } = "";

    public bool IsDynamic => RecipientTypeDetails.Contains("Dynamic", StringComparison.OrdinalIgnoreCase);
    public bool IsUnified => RecipientTypeDetails.Contains("GroupMailbox", StringComparison.OrdinalIgnoreCase);
}
