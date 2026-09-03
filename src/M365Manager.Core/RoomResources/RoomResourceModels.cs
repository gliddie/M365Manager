namespace M365Manager.Core.RoomResources;

/// <summary>Room mailbox vs. equipment/resource mailbox (legacy: the chkNewGRRoom / chkNewGREquip radio pair).</summary>
public enum ResourceKind
{
    Room,
    Equipment,
}

/// <summary>
/// Who may book the room and who approves out-of-policy requests. Replaces the legacy tool's four
/// interdependent checkboxes (chkStdDele / chkGenUseCustDele / chkResDele / chkResUsr), whose
/// precedence rules were implicit and scattered across RoomEquipGroups.ps1 and RoomEquipConfig.ps1.
/// </summary>
public enum RoomAccessModel
{
    /// <summary>General use. Anyone may book in-policy; the shared site delegate group (MBX.&lt;SITE&gt;.RRS.OutOfPolicy.DE) approves out-of-policy requests.</summary>
    GeneralUseSiteDelegates,

    /// <summary>General use, but with its own delegate group (MBX.&lt;SITE&gt;.RRS.OutOfPolicy.&lt;Room&gt;.DE) instead of the site-wide one.</summary>
    GeneralUseCustomDelegates,

    /// <summary>Restricted. Only members of the room's users group (…&lt;Room&gt;.US) may book at all; its delegate group approves.</summary>
    Restricted,
}

/// <summary>
/// Calendar booking rules. The legacy tool had four near-duplicate config scripts
/// (ConfigConferenceRoom / ConfigRestrictedConferenceRoom / ConfigNonRecurringConferenceRoom /
/// ConfigRestrictedNonRecurringConferenceRoom) that crossed this dimension with
/// <see cref="RoomAccessModel"/>; here the two are orthogonal.
/// </summary>
public enum BookingPolicyPreset
{
    /// <summary>Recurring meetings allowed, 90-day window, 24h max, limited conflicts (3 instances / 20%).</summary>
    Standard,

    /// <summary>Hoteling / hot-desk: no recurring meetings, no conflicts at all (legacy "NonRecurring" variants).</summary>
    Hoteling,

    /// <summary>Values taken from the form instead of a preset.</summary>
    Custom,
}

/// <summary>The concrete Set-CalendarProcessing values behind a <see cref="BookingPolicyPreset"/>.</summary>
public sealed record BookingPolicy(
    int BookingWindowInDays,
    int MaximumDurationInMinutes,
    bool AllowRecurringMeetings,
    int MaximumConflictInstances,
    int ConflictPercentageAllowed)
{
    /// <summary>Legacy standard policy: 90 days / 1440 minutes / recurring allowed / 3 conflicts / 20%.</summary>
    public static BookingPolicy Standard { get; } = new(90, 1440, true, 3, 20);

    /// <summary>Legacy non-recurring ("hoteling") policy: same window, no recurrence, zero conflicts.</summary>
    public static BookingPolicy Hoteling { get; } = new(90, 1440, false, 0, 0);

    public static BookingPolicy For(BookingPolicyPreset preset) => preset switch
    {
        BookingPolicyPreset.Standard => Standard,
        BookingPolicyPreset.Hoteling => Hoteling,
        _ => Standard,
    };
}

/// <summary>Which of a room's two groups a membership change targets.</summary>
public enum RoomGroupKind
{
    /// <summary>The delegate group (.DE) - approves out-of-policy requests, has FullAccess on the mailbox.</summary>
    Delegates,

    /// <summary>The authorized-users group (.US) - restricted rooms only; only its members may book.</summary>
    Users,
}

public enum MembershipChangeMode
{
    Add,
    Remove,
}
