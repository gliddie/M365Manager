using M365Manager.Core.Exchange;

namespace M365Manager.Core.RoomResources;

/// <summary>
/// Input for <see cref="IRoomResourceService.CreateAsync"/> - mirrors the fields the legacy
/// "New Conference Room / New Equipment-Resource" WinForms dialog (RoomResourceNewForm.ps1) asked
/// for, with the four interdependent delegate checkboxes replaced by <see cref="AccessModel"/>.
/// </summary>
public sealed class RoomResourceCreationRequest
{
    /// <summary>ServiceNow (or equivalent) task number authorizing the creation. Required.</summary>
    public required string TaskNumber { get; init; }

    /// <summary>
    /// Who asked for the room - the confirmation e-mail goes to them. A room has no owner field the
    /// way a shared mailbox does, and the legacy templates were addressed by hand by the operator,
    /// so this is the equivalent. Empty means no confirmation e-mail is sent.
    /// </summary>
    public string RequesterIdentity { get; init; } = "";

    /// <summary>"&lt;SITE&gt; &lt;Name&gt;" as typed, e.g. "NBK Meeting Room 1". The first word is the site code.</summary>
    public required string RawName { get; init; }

    public required ResourceKind Kind { get; init; }
    public required RoomAccessModel AccessModel { get; init; }
    public required BookingPolicyPreset Preset { get; init; }

    /// <summary>Used when <see cref="Preset"/> is <see cref="BookingPolicyPreset.Custom"/>; otherwise the preset's values win.</summary>
    public BookingPolicy? CustomPolicy { get; init; }

    /// <summary>Rooms only. Part of the address and set as ResourceCapacity.</summary>
    public string Capacity { get; init; } = "";

    public string Building { get; init; } = "";

    /// <summary>Floor number, or "0"/"Ground" for the ground floor.</summary>
    public string Floor { get; init; } = "";

    /// <summary>Comma-separated identities to add to the delegate group (only used when the group is created here).</summary>
    public string DelegateMembers { get; init; } = "";

    /// <summary>Comma-separated identities to add to the users group. Restricted rooms only.</summary>
    public string UserMembers { get; init; } = "";

    /// <summary>Overrides Settings' room mailbox domain when non-empty.</summary>
    public string MailboxDomain { get; init; } = "";

    /// <summary>Overrides Settings' room group domain when non-empty.</summary>
    public string GroupDomain { get; init; } = "";
}

/// <summary>
/// What a create would do, computed without writing anything - the pre-flight check the legacy
/// tool lacked (it discovered collisions halfway through creation).
/// </summary>
public sealed class RoomCreationPreview
{
    public required RoomNames Names { get; init; }

    /// <summary>Non-empty when the site code has no dbo.RoomSites entry - creation would fail.</summary>
    public string? UnknownSiteCode { get; init; }

    public string? TimeZone { get; init; }
    public string? RegionalAdminGroup { get; init; }

    // Existence probes: true = exists, false = does not exist, null = could not be determined
    // (Exchange not connected, transient failure) or not applicable to this request.
    public bool? MailboxExists { get; init; }
    public bool? RoomListExists { get; init; }
    public bool? DelegateGroupExists { get; init; }
    public bool? UsersGroupExists { get; init; }

    /// <summary>Blocking problems. Empty means creation can proceed.</summary>
    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    public bool CanCreate => Errors.Count == 0;

    /// <summary>
    /// True when a probe that creation depends on came back inconclusive. Creating anyway risks
    /// aborting partway - e.g. trying to create a room list that already exists, after the mailbox
    /// has already been made. The live preview tolerates this; <see cref="IRoomResourceService.CreateAsync"/> does not.
    /// </summary>
    public bool HasInconclusiveProbes =>
        MailboxExists is null || DelegateGroupExists is null
        || (RoomListRequired && RoomListExists is null)
        || (UsersGroupRequired && UsersGroupExists is null);

    /// <summary>Set by the service: whether this request needs a room list / users group at all.</summary>
    public bool RoomListRequired { get; init; }
    public bool UsersGroupRequired { get; init; }
}

public sealed class RoomResourceCreationResult
{
    public required bool Succeeded { get; init; }
    public string? DisplayName { get; init; }
    public string? PrimarySmtpAddress { get; init; }
    public string? ErrorMessage { get; init; }

    /// <summary>Set when the operation succeeded but the confirmation e-mail did not go out.</summary>
    public string? WarningMessage { get; init; }

    public Guid CorrelationId { get; init; }

    public static RoomResourceCreationResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}

/// <summary>
/// Input for <see cref="IRoomResourceService.UpdateDetailsAsync"/>. Covers what the legacy
/// "Rename Room and Change Room Details" dialog changed apart from the name/address itself -
/// renaming is a separate operation and not implemented yet.
/// </summary>
public sealed class RoomDetailsUpdateRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>Who asked for the change; the confirmation e-mail goes to them. Empty sends none.</summary>
    public string RequesterIdentity { get; init; } = "";

    /// <summary>Primary SMTP address of the room/equipment mailbox.</summary>
    public required string Address { get; init; }

    public required ResourceKind Kind { get; init; }

    /// <summary>Rooms only. Empty leaves the current capacity untouched.</summary>
    public string Capacity { get; init; } = "";

    /// <summary>Empty leaves the current time zone untouched.</summary>
    public string TimeZone { get; init; } = "";

    /// <summary>When null the calendar booking rules are left as they are.</summary>
    public BookingPolicyPreset? Preset { get; init; }

    public BookingPolicy? CustomPolicy { get; init; }
}

/// <summary>
/// Input for <see cref="IRoomResourceService.ChangeMembershipAsync"/> - the
/// RoomResourceOOPChanges.ps1 equivalent.
/// </summary>
public sealed class RoomMembershipChangeRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>Who asked for the change; the confirmation e-mail goes to them. Empty sends none.</summary>
    public string RequesterIdentity { get; init; } = "";

    /// <summary>Primary SMTP address of the room/equipment mailbox (used for the audit trail).</summary>
    public required string Address { get; init; }

    /// <summary>The group to change - resolved from the cached row, so the operator never types a group name.</summary>
    public required string GroupIdentity { get; init; }

    public required RoomGroupKind GroupKind { get; init; }
    public required MembershipChangeMode Mode { get; init; }

    /// <summary>Comma-separated identities (SamAccountName, UPN, e-mail, or a group name).</summary>
    public required string Identities { get; init; }
}

public sealed class RoomMembershipChangeResult
{
    public required bool Succeeded { get; init; }
    public IReadOnlyList<MemberOperationResult> Results { get; init; } = Array.Empty<MemberOperationResult>();
    public string? ErrorMessage { get; init; }

    /// <summary>Set when the change went through but the confirmation e-mail did not go out.</summary>
    public string? WarningMessage { get; init; }

    public Guid CorrelationId { get; init; }

    public static RoomMembershipChangeResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}

/// <summary>Input for <see cref="IRoomResourceService.RemoveAsync"/> - the RoomResourceRemove.ps1 equivalent.</summary>
public sealed class RoomRemovalRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>Who asked for the removal; the confirmation e-mail goes to them. Empty sends none.</summary>
    public string RequesterIdentity { get; init; } = "";

    /// <summary>Primary SMTP address of the room/equipment mailbox to delete.</summary>
    public required string Address { get; init; }

    /// <summary>
    /// Also delete the room's own delegate/users groups and drop it from its room list. Legacy left
    /// these behind as orphans. Site-wide delegate groups are never deleted, only per-room ones.
    /// </summary>
    public bool RemoveOrphanedGroups { get; init; } = true;
}

public sealed class RoomRemovalResult
{
    public required bool Succeeded { get; init; }

    /// <summary>
    /// The full pre-removal snapshot (mailbox, statistics, folder statistics, permissions, calendar
    /// processing/configuration/regional configuration) that the legacy script wrote to a text file
    /// before deleting. Also written to the audit log under <see cref="CorrelationId"/>.
    /// </summary>
    public string Snapshot { get; init; } = "";

    public IReadOnlyList<string> RemovedGroups { get; init; } = Array.Empty<string>();
    public string? ErrorMessage { get; init; }

    /// <summary>Set when the removal went through but the confirmation e-mail did not go out.</summary>
    public string? WarningMessage { get; init; }

    public Guid CorrelationId { get; init; }

    public static RoomRemovalResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}

/// <summary>One row of the Rooms &amp; Resources overview grid - read from the SQL cache (dbo.RoomResources), never live from M365.</summary>
public sealed class RoomResourceOverviewRow
{
    public string ExchangeGuid { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Alias { get; init; } = "";
    public string PrimarySmtpAddress { get; init; } = "";
    public string ResourceKind { get; init; } = "";
    public int? Capacity { get; init; }
    public string Building { get; init; } = "";
    public string Floor { get; init; } = "";
    public string Office { get; init; } = "";
    public string SiteCode { get; init; } = "";
    public string TimeZone { get; init; } = "";
    public string RoomListName { get; init; } = "";
    public string DelegateGroup { get; init; } = "";
    public string UsersGroup { get; init; } = "";
    public string AccessModel { get; init; } = "";
    public string BookingPolicy { get; init; } = "";
    public int? BookingWindowInDays { get; init; }
    public int? MaximumDurationInMinutes { get; init; }
    public bool? AllowRecurringMeetings { get; init; }
    public bool IsDeletedInM365 { get; init; }
    public DateTime? CreatedDateTime { get; init; }
    public DateTime LastImportedAtUtc { get; init; }

    /// <summary>True for restricted rooms, which are the only ones with an authorized-users group.</summary>
    public bool HasUsersGroup => !string.IsNullOrWhiteSpace(UsersGroup);
}
