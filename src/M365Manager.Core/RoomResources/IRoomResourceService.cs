namespace M365Manager.Core.RoomResources;

/// <summary>
/// Orchestrates room and equipment mailboxes the way the legacy Room &amp; Resource Admin Menu did
/// (RoomResourceNewForm/RoomEquipNew/RoomEquipGroups/RoomEquipConfig, RoomResourceOOPChanges,
/// RoomResourceRemove): enforced naming, room list + delegate/users group provisioning, calendar
/// booking policy, mailbox and calendar permissions, time zone, renaming (RoomResourceRename.ps1),
/// and a full step-by-step audit trail in dbo.LogEntries.
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

    /// <summary>
    /// Reads a room's delegates and the people allowed to book it, live from Exchange - whether the
    /// calendar points at a group or lists people directly. Throws with an operator-readable message
    /// when the room cannot be read.
    /// </summary>
    Task<RoomAccessSnapshot> GetAccessAsync(string roomAddress, CancellationToken ct = default);

    /// <summary>
    /// Resolves every typed name to exactly one recipient and works out what an access change
    /// (RoomResourceOOPChanges.ps1) would add and remove. Writes nothing.
    /// </summary>
    Task<RoomAccessChangePlan> PlanAccessChangeAsync(RoomAccessChangeRequest request, CancellationToken ct = default);

    /// <summary>Applies a plan from <see cref="PlanAccessChangeAsync"/>: both lists, one audit trail, one confirmation mail.</summary>
    Task<RoomAccessChangeResult> ApplyAccessChangeAsync(RoomAccessChangePlan plan, CancellationToken ct = default);

    /// <summary>Captures a full snapshot into the audit log, then deletes the mailbox (and optionally its orphaned groups).</summary>
    Task<RoomRemovalResult> RemoveAsync(RoomRemovalRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>
    /// Reads a room back for renaming: names, address, access model, groups, time zone. Throws
    /// when it cannot be resolved or is not a room/equipment mailbox; the message is for the operator.
    /// </summary>
    Task<RoomRenameLookup> LookUpForRenameAsync(string mailboxIdentity, CancellationToken ct = default);

    /// <summary>Works out what renaming <paramref name="current"/> to the request's name would change. Writes nothing.</summary>
    Task<RoomRenamePlan> PlanRenameAsync(RoomRenameLookup current, RoomRenameRequest request, CancellationToken ct = default);

    /// <summary>
    /// Renames the mailbox (old address kept as an alias), moves the UPN along, renames the room's
    /// own groups and - when the site code changes - moves it to the new site's room list,
    /// delegates, regional admins and time zone.
    /// </summary>
    Task<RoomRenameResult> RenameAsync(RoomRenameRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>
    /// After telephony was set up on the Telephony tab: mails the requester the room's details with
    /// its new number. Returns a warning for the UI, or null when sent (or no requester was given).
    /// </summary>
    Task<string?> SendTelephonyReadyMailAsync(RoomResourceOverviewRow room, string requesterIdentity, string taskNumber,
        string phoneNumber, bool mentionPassword = false, CancellationToken ct = default);

    /// <summary>
    /// Enables a room's account with a new random password that never expires - for a room whose
    /// account was never enabled, or a Teams Rooms device that needs a new password. The password
    /// comes back once for the operator; it is never logged, stored or mailed.
    /// </summary>
    Task<RoomPasswordResult> SetNewAccountPasswordAsync(string roomAddress, string taskNumber, CancellationToken ct = default);

    /// <summary>Reads the overview grid from the SQL cache (dbo.RoomResources) - never live from M365.</summary>
    Task<IReadOnlyList<RoomResourceOverviewRow>> GetOverviewAsync(CancellationToken ct = default);
}
