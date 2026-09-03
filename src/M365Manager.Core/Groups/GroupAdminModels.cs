using M365Manager.Core.Exchange;

namespace M365Manager.Core.Groups;

/// <summary>What <c>New-DistributionGroup -Type</c> is given.</summary>
public enum GroupKind
{
    /// <summary>Plain distribution list.</summary>
    Distribution,

    /// <summary>Mail-enabled security group - can also hold permissions.</summary>
    Security,
}

/// <summary>How identities are applied to an existing list (owners, authorized senders, aliases).</summary>
public enum ListChangeMode
{
    Add,
    Remove,
    Replace,
}

/// <summary>
/// Input for <see cref="IGroupAdminService.CreateAsync"/> - the fields of the legacy two-step
/// NewDLForm dialog (Build-DLInputForm followed by DefaultForm).
/// </summary>
public sealed class GroupCreationRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>
    /// Who asked for the group - the confirmation e-mail goes to them, as the legacy DLNew.oft was
    /// addressed by hand ("Per your request ..."). Empty means no e-mail is sent.
    /// </summary>
    public string RequesterIdentity { get; init; } = "";

    /// <summary>Dot-notation name as typed, e.g. "LST.EMEA.Some Team".</summary>
    public required string RawName { get; init; }

    public required GroupKind Kind { get; init; }
    public required string Domain { get; init; }

    /// <summary>Inserts "TMP" as the second segment and records a removal date in the Notes.</summary>
    public bool IsTemporary { get; init; }

    /// <summary>Only meaningful for a temporary group; ends up in the group's Notes.</summary>
    public DateTime? RemoveOn { get; init; }

    /// <summary>Single character stripped out of the third name segment (legacy "Remove Special Character").</summary>
    public string RemoveCharacter { get; init; } = "";

    /// <summary>Comma-separated identities set as ManagedBy.</summary>
    public required string Owners { get; init; }

    /// <summary>Comma-separated identities added as members. May include other lists (LST.*/DST.*).</summary>
    public string Members { get; init; } = "";

    /// <summary>Comma-separated identities allowed to send to the group (AcceptMessagesOnlyFromSendersOrMembers).</summary>
    public string AuthorizedSenders { get; init; } = "";

    /// <summary>Comma-separated extra SMTP addresses added alongside the primary one.</summary>
    public string AdditionalAliases { get; init; } = "";

    /// <summary>False sets RequireSenderAuthenticationEnabled, i.e. internal senders only.</summary>
    public bool AllowExternalSenders { get; init; }

    /// <summary>Overrides the derived display name when non-empty.</summary>
    public string DisplayNameOverride { get; init; } = "";

    /// <summary>Overrides the derived primary SMTP address when non-empty.</summary>
    public string AddressOverride { get; init; } = "";
}

/// <summary>
/// Input for <see cref="IGroupAdminService.CreateDynamicAsync"/>. The legacy script hardcoded one
/// filter per site type; these are the same conditions, exposed as fields.
/// </summary>
public sealed class DynamicGroupCreationRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>Who asked for the group; the confirmation e-mail goes to them. Empty sends none.</summary>
    public string RequesterIdentity { get; init; } = "";
    public required string DisplayName { get; init; }
    public required string Domain { get; init; }

    /// <summary>Overrides the address derived from the display name when non-empty.</summary>
    public string AddressOverride { get; init; } = "";

    /// <summary>Graph/Exchange recipient filter, e.g. "MailboxUsers".</summary>
    public string IncludedRecipients { get; init; } = "MailboxUsers";

    public string ConditionalCompany { get; init; } = "";
    public string ConditionalCustomAttribute1 { get; init; } = "";
    public string ConditionalCustomAttribute2 { get; init; } = "";
    public string ConditionalCustomAttribute3 { get; init; } = "";
    public string ConditionalCustomAttribute8 { get; init; } = "";

    public string Notes { get; init; } = "";
    public string MailTip { get; init; } = "";
}

/// <summary>What a create would do, computed without writing anything.</summary>
public sealed class GroupCreationPreview
{
    public required GroupNames Names { get; init; }

    // true = exists, false = does not, null = could not be determined.
    public bool? GroupExists { get; init; }

    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    public bool CanCreate => Errors.Count == 0;

    /// <summary>Creation must not proceed on an inconclusive probe - see RoomResourceService for the same rule.</summary>
    public bool HasInconclusiveProbes => GroupExists is null;
}

public sealed class GroupOperationResult
{
    public required bool Succeeded { get; init; }
    public string? DisplayName { get; init; }
    public string? Address { get; init; }
    public string? ErrorMessage { get; init; }
    public Guid CorrelationId { get; init; }

    /// <summary>Per-identity outcomes for the operations that take a list.</summary>
    public IReadOnlyList<MemberOperationResult> Results { get; init; } = Array.Empty<MemberOperationResult>();

    /// <summary>Set when the group was created/changed but a follow-up step failed.</summary>
    public string? Warning { get; init; }

    /// <summary>Result detail worth showing back, e.g. the MailTip a refresh produced.</summary>
    public string? Info { get; init; }

    public static GroupOperationResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}

/// <summary>Input for the operations that add/remove/replace a list of identities on an existing group.</summary>
public sealed class GroupListChangeRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>
    /// Who asked for the change; the confirmation e-mail goes to them. Empty sends none. Alias
    /// changes send nothing regardless - there is no legacy template for them.
    /// </summary>
    public string RequesterIdentity { get; init; } = "";

    /// <summary>Identity of the group - its primary SMTP address or name.</summary>
    public required string GroupIdentity { get; init; }

    public required ListChangeMode Mode { get; init; }

    /// <summary>Comma-separated identities.</summary>
    public required string Identities { get; init; }

    /// <summary>Set for Microsoft 365 groups, which use the UnifiedGroupLinks cmdlets for owners.</summary>
    public bool IsUnifiedGroup { get; init; }
}

/// <summary>Input for <see cref="IGroupAdminService.RenameAsync"/>.</summary>
public sealed class GroupRenameRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>Who asked for the rename; the confirmation e-mail goes to them. Empty sends none.</summary>
    public string RequesterIdentity { get; init; } = "";

    /// <summary>The group as it is now.</summary>
    public required string GroupIdentity { get; init; }

    /// <summary>New display name / name, dot-notation as typed.</summary>
    public required string NewName { get; init; }

    /// <summary>
    /// New primary SMTP address. The old one is kept as a secondary alias, as the legacy script did,
    /// so mail to the previous address keeps working.
    /// </summary>
    public required string NewAddress { get; init; }
}

/// <summary>Input for <see cref="IGroupAdminService.RemoveAsync"/>.</summary>
public sealed class GroupRemovalRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>Who asked for the removal; the confirmation e-mail goes to them. Empty sends none.</summary>
    public string RequesterIdentity { get; init; } = "";
    public required string GroupIdentity { get; init; }
}

public sealed class GroupRemovalResult
{
    public required bool Succeeded { get; init; }

    /// <summary>Members/owners/settings captured before deletion, also written to the audit log.</summary>
    public string Snapshot { get; init; } = "";

    /// <summary>Which cmdlet path was taken: Distribution, Dynamic or Unified.</summary>
    public string RemovedAs { get; init; } = "";

    public string? ErrorMessage { get; init; }

    /// <summary>Set when the group was removed but the confirmation e-mail did not go out.</summary>
    public string? Warning { get; init; }

    public Guid CorrelationId { get; init; }

    public static GroupRemovalResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}
