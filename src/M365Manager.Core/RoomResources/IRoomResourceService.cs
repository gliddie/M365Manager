namespace M365Manager.Core.RoomResources;

/// <summary>
/// Orchestrates room and equipment mailboxes the way the legacy Room &amp; Resource Admin Menu did
/// (RoomResourceNewForm/RoomEquipNew/RoomEquipGroups/RoomEquipConfig, RoomResourceOOPChanges,
/// RoomResourceRemove): enforced naming, room list + delegate/users group provisioning, calendar
/// booking policy, mailbox and calendar permissions, time zone, and a full step-by-step audit
/// trail in dbo.LogEntries.
///
/// Renaming a room (legacy RoomResourceRename.ps1) is deliberately not implemented yet - it also
/// changes the UPN via Graph and renames the associated groups, and is scheduled as its own pass.
/// </summary>
public interface IRoomResourceService
{
    /// <summary>
    /// Resolves every derived name and checks what already exists, without writing anything.
    /// Drives the form's live preview and gates <see cref="CreateAsync"/>.
    /// </summary>
    Task<RoomCreationPreview> PreviewAsync(RoomResourceCreationRequest request, CancellationToken ct = default);

    Task<RoomResourceCreationResult> CreateAsync(RoomResourceCreationRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>Changes capacity/location/time zone/booking policy on an existing room. Does not touch its name or address.</summary>
    Task<RoomResourceCreationResult> UpdateDetailsAsync(RoomDetailsUpdateRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>Adds or removes members of a room's delegate or authorized-users group.</summary>
    Task<RoomMembershipChangeResult> ChangeMembershipAsync(RoomMembershipChangeRequest request, CancellationToken ct = default);

    /// <summary>Reads the current members of one of a room's groups, for the management panel.</summary>
    Task<IReadOnlyList<string>> GetGroupMembersAsync(string groupIdentity, CancellationToken ct = default);

    /// <summary>Captures a full snapshot into the audit log, then deletes the mailbox (and optionally its orphaned groups).</summary>
    Task<RoomRemovalResult> RemoveAsync(RoomRemovalRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>Reads the overview grid from the SQL cache (dbo.RoomResources) - never live from M365.</summary>
    Task<IReadOnlyList<RoomResourceOverviewRow>> GetOverviewAsync(CancellationToken ct = default);
}
