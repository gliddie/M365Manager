namespace M365Manager.Core.RoomResources;

/// <summary>The two lists an "Access Change to Restricted Room" ticket names.</summary>
public enum RoomAccessListKind
{
    /// <summary>ResourceDelegates - approve out-of-policy requests.</summary>
    Delegates,

    /// <summary>BookInPolicy - on a restricted room, the only people who may book it.</summary>
    BookingUsers,
}

/// <summary>Where a list's members live.</summary>
public enum RoomAccessStorage
{
    /// <summary>The calendar points at one group (.DE / .US); access changes are group membership changes.</summary>
    Group,

    /// <summary>
    /// The calendar lists people directly - how older restricted rooms were set up when no group
    /// was created. Access changes rewrite the calendar setting itself.
    /// </summary>
    CalendarEntries,
}

/// <summary>One person (or group) on a room's delegate or booking list.</summary>
/// <param name="Identity">What to hand Exchange for this member - its address, or the raw entry when it has none.</param>
/// <param name="ObjectGuid">Exchange Guid, for comparing what was typed with who is already there.</param>
public sealed record RoomAccessMember(string Identity, string DisplayName, string Address, Guid? ObjectGuid)
{
    public string Label => Address.Length > 0 && !string.Equals(Address, DisplayName, StringComparison.OrdinalIgnoreCase)
        ? $"{DisplayName} <{Address}>"
        : DisplayName;
}

public sealed class RoomAccessList
{
    public required RoomAccessListKind Kind { get; init; }
    public required RoomAccessStorage Storage { get; init; }

    /// <summary>Group the calendar points at (its address); empty for <see cref="RoomAccessStorage.CalendarEntries"/>.</summary>
    public string GroupIdentity { get; init; } = "";

    /// <summary>The group's name as admins know it, e.g. "MBX.MIL.Black.DE".</summary>
    public string GroupName { get; init; } = "";

    /// <summary>
    /// A site-wide delegate group (MBX.&lt;SITE&gt;.RRS.OutOfPolicy.DE) is shared by every general-use
    /// room at the site - changing it changes all of them.
    /// </summary>
    public bool IsSiteWide { get; init; }

    public IReadOnlyList<RoomAccessMember> Members { get; init; } = Array.Empty<RoomAccessMember>();

    /// <summary>"MBX.MIL.Black.DE" or "set directly on the calendar", for the UI and the confirmation.</summary>
    public string Source => Storage == RoomAccessStorage.Group
        ? (GroupName.Length > 0 ? GroupName : GroupIdentity)
        : "set directly on the calendar";
}

/// <summary>A room's delegate and booking lists, read live from Exchange - never from the SQL cache.</summary>
public sealed class RoomAccessSnapshot
{
    public required string DisplayName { get; init; }
    public required string Address { get; init; }
    public required RoomAccessModel AccessModel { get; init; }
    public required RoomAccessList Delegates { get; init; }

    /// <summary>Null on a general-use room: anyone may book it, there is no list to change.</summary>
    public RoomAccessList? BookingUsers { get; init; }

    /// <summary>
    /// Booking rights come from the delegate group itself, so a change to one list is a change to
    /// both. Happens on restricted rooms whose BookInPolicy was pointed at the .DE group.
    /// </summary>
    public bool BookingSharesDelegateGroup =>
        BookingUsers is { Storage: RoomAccessStorage.Group } users
        && Delegates.Storage == RoomAccessStorage.Group
        && string.Equals(users.GroupIdentity, Delegates.GroupIdentity, StringComparison.OrdinalIgnoreCase);
}

/// <summary>
/// The ticket as typed - "names of individuals to be added or removed" for each of the two lists.
/// Each field takes one identity per line, or several separated by ";" or ",".
/// </summary>
public sealed class RoomAccessChangeRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>Who asked for the change; the confirmation e-mail goes to them. Empty sends none.</summary>
    public string RequesterIdentity { get; init; } = "";

    /// <summary>Primary SMTP address of the room/equipment mailbox.</summary>
    public required string Address { get; init; }

    public string AddDelegates { get; init; } = "";
    public string RemoveDelegates { get; init; } = "";
    public string AddBookingUsers { get; init; } = "";
    public string RemoveBookingUsers { get; init; } = "";

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(AddDelegates) && string.IsNullOrWhiteSpace(RemoveDelegates)
        && string.IsNullOrWhiteSpace(AddBookingUsers) && string.IsNullOrWhiteSpace(RemoveBookingUsers);
}

public enum RoomAccessChangeOutcome
{
    WillAdd,
    WillRemove,

    /// <summary>Asked to add someone who is already on the list - nothing to do.</summary>
    AlreadyOnList,

    /// <summary>Asked to remove someone who is not on the list - nothing to do.</summary>
    NotOnList,

    NotFound,

    /// <summary>The typed text matches more than one recipient; nobody is guessed.</summary>
    Ambiguous,

    /// <summary>The same person is to be added to and removed from the same list.</summary>
    Conflict,
}

public static class RoomAccessChangeOutcomeText
{
    /// <summary>Why an entry is skipped, for the confirmation, the status line and the log.</summary>
    public static string Describe(this RoomAccessChangeOutcome outcome) => outcome switch
    {
        RoomAccessChangeOutcome.AlreadyOnList => "already on the list",
        RoomAccessChangeOutcome.NotOnList => "not on the list",
        RoomAccessChangeOutcome.NotFound => "no recipient found",
        RoomAccessChangeOutcome.Ambiguous => "matches several recipients - use the e-mail address",
        RoomAccessChangeOutcome.Conflict => "both added and removed",
        _ => "",
    };
}

public sealed record RoomAccessChangeItem(
    RoomAccessListKind List,
    bool IsAdd,
    string Typed,
    RoomAccessMember? Member,
    RoomAccessChangeOutcome Outcome)
{
    public bool IsWork => Outcome is RoomAccessChangeOutcome.WillAdd or RoomAccessChangeOutcome.WillRemove;

    /// <summary>Resolved name when there is one, otherwise what was typed.</summary>
    public string Label => Member?.Label ?? Typed;
}

/// <summary>What applying a request would do. Built by resolving everything first, writing nothing.</summary>
public sealed class RoomAccessChangePlan
{
    public required RoomAccessSnapshot Room { get; init; }
    public required RoomAccessChangeRequest Request { get; init; }
    public IReadOnlyList<RoomAccessChangeItem> Items { get; init; } = Array.Empty<RoomAccessChangeItem>();

    /// <summary>Reasons the request cannot be applied as typed. Non-empty blocks the apply.</summary>
    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    /// <summary>Consequences worth a second look - shown in the confirmation, not blocking.</summary>
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();

    public bool HasWork => Items.Any(i => i.IsWork);
    public bool HasRemovals => Items.Any(i => i.Outcome == RoomAccessChangeOutcome.WillRemove);
    public IEnumerable<RoomAccessChangeItem> Skipped => Items.Where(i => !i.IsWork);
}

public sealed record RoomAccessChangeItemResult(RoomAccessChangeItem Item, bool Succeeded, string? Error);

public sealed class RoomAccessChangeResult
{
    public required bool Succeeded { get; init; }
    public IReadOnlyList<RoomAccessChangeItemResult> Results { get; init; } = Array.Empty<RoomAccessChangeItemResult>();
    public string? ErrorMessage { get; init; }

    /// <summary>Set when the change went through but the confirmation e-mail did not go out.</summary>
    public string? WarningMessage { get; init; }

    public Guid CorrelationId { get; init; }

    public static RoomAccessChangeResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}
