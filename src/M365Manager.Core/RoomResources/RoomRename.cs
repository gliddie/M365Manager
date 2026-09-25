namespace M365Manager.Core.RoomResources;

/// <summary>
/// A room or equipment mailbox as it is now, read back for the rename form - the legacy
/// RoomResourceRename.ps1 "Get Room Details" stage.
/// </summary>
public sealed class RoomRenameLookup
{
    public required Guid ExchangeGuid { get; init; }
    public required string DisplayName { get; init; }
    public required string PrimarySmtpAddress { get; init; }
    public required string Alias { get; init; }
    public required ResourceKind Kind { get; init; }
    public required RoomAccessModel AccessModel { get; init; }

    /// <summary>First word of the display name, upper-cased - the site code everything else hangs off.</summary>
    public required string SiteCode { get; init; }

    public int? Capacity { get; init; }

    /// <summary>Parsed back out of the address ("FRKBldg1.FLR2..."), to prefill the form. Empty when absent.</summary>
    public string Building { get; init; } = "";
    public string Floor { get; init; } = "";

    public string TimeZone { get; init; } = "";

    /// <summary>ResourceDelegates on the calendar - the delegate group(s).</summary>
    public IReadOnlyList<string> DelegateGroups { get; init; } = Array.Empty<string>();

    /// <summary>BookInPolicy on the calendar - the authorized-users group(s) of a restricted room.</summary>
    public IReadOnlyList<string> UsersGroups { get; init; } = Array.Empty<string>();

    /// <summary>Entra object id, for moving the UPN along with the address.</summary>
    public string ExternalDirectoryObjectId { get; init; } = "";
    public string UserPrincipalName { get; init; } = "";
}

public sealed class RoomRenameRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>Who asked for the rename; the confirmation e-mail goes to them. Empty sends none.</summary>
    public string RequesterIdentity { get; init; } = "";

    /// <summary>The room as it is now - address, alias or display name.</summary>
    public required string MailboxIdentity { get; init; }

    /// <summary>New "&lt;SITE&gt; &lt;Name&gt;", as on the create form.</summary>
    public required string RawName { get; init; }

    public string Building { get; init; } = "";
    public string Floor { get; init; } = "";

    /// <summary>Rooms only.</summary>
    public string Capacity { get; init; } = "";
}

/// <summary>What a rename would do, worked out without changing anything - the form's preview.</summary>
public sealed class RoomRenamePlan
{
    public required RoomRenameLookup Current { get; init; }
    public required RoomNames NewNames { get; init; }

    public bool SiteChanged { get; init; }

    /// <summary>Per-room delegate/users groups whose name follows the room's: (current, new).</summary>
    public IReadOnlyList<(string From, string To)> GroupRenames { get; init; } = Array.Empty<(string, string)>();

    /// <summary>Site-wide delegate group swapped for the new site's, general-use rooms moving site: (old, new).</summary>
    public (string From, string To)? DelegateSwap { get; init; }

    /// <summary>Room list move, rooms moving site: (old, new).</summary>
    public (string From, string To)? RoomListMove { get; init; }

    /// <summary>Set when the new site's time zone differs from the room's current one.</summary>
    public string? NewTimeZone { get; init; }

    /// <summary>The new site's regional RRS admin group, when the site changes.</summary>
    public string? NewRegionalAdminGroup { get; init; }

    /// <summary>Anything that blocks the rename. Empty means it can go ahead.</summary>
    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    /// <summary>Things worth knowing that do not block it.</summary>
    public IReadOnlyList<string> Notes { get; init; } = Array.Empty<string>();
}

public sealed class RoomRenameResult
{
    public required bool Succeeded { get; init; }

    public string? PreviousDisplayName { get; init; }
    public string? PreviousPrimarySmtpAddress { get; init; }
    public string? DisplayName { get; init; }
    public string? PrimarySmtpAddress { get; init; }

    /// <summary>One line per follow-up change: groups, room list, time zone, admins, UPN.</summary>
    public IReadOnlyList<string> Changes { get; init; } = Array.Empty<string>();

    public string? ErrorMessage { get; init; }

    /// <summary>The mailbox was renamed, but a follow-up step failed. <see cref="Succeeded"/> stays true.</summary>
    public string? WarningMessage { get; init; }

    public Guid CorrelationId { get; init; }

    public static RoomRenameResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}
