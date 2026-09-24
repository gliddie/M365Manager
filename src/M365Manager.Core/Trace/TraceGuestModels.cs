namespace M365Manager.Core.Trace;

/// <summary>
/// Input for <see cref="ITraceGuestService.InviteAsync"/> - the two parameters the legacy invite
/// script took, plus an optional ticket number for the audit trail.
/// </summary>
public sealed class TraceGuestInvitationRequest
{
    /// <summary>Name the guest is invited under, e.g. "Ralf Terjung".</summary>
    public required string DisplayName { get; init; }

    /// <summary>The customer's own e-mail address - where the invitation is sent.</summary>
    public required string Email { get; init; }

    /// <summary>Optional; recorded in the log when given. Guest invitations don't always come from a ticket.</summary>
    public string TaskNumber { get; init; } = "";
}

public sealed class TraceGuestInvitationResult
{
    public required bool Succeeded { get; init; }
    public string? DisplayName { get; init; }
    public string? Email { get; init; }

    /// <summary>Entra object id of the invited guest, straight from the invitation response.</summary>
    public string? UserId { get; init; }

    /// <summary>The redemption URL Graph generated. Useful when the invitation mail doesn't arrive.</summary>
    public string? RedeemUrl { get; init; }

    /// <summary>Set when the guest was invited but stamping the company name afterwards failed.</summary>
    public string? Warning { get; init; }

    public string? ErrorMessage { get; init; }
    public Guid CorrelationId { get; init; }

    public static TraceGuestInvitationResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}

/// <summary>One row of the Trace guest overview - read live from Entra, not from a SQL cache.</summary>
public sealed class TraceGuestRow
{
    public string Id { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string Mail { get; init; } = "";
    public string UserPrincipalName { get; init; } = "";
    public string CompanyName { get; init; } = "";

    /// <summary>Graph's externalUserState: "Accepted" or "PendingAcceptance".</summary>
    public string InvitationState { get; init; } = "";

    public DateTime? StateChangedUtc { get; init; }
    public DateTime? CreatedUtc { get; init; }
    public bool AccountEnabled { get; init; }

    /// <summary>Friendly label for the grid.</summary>
    public string StatusLabel => InvitationState switch
    {
        "Accepted" => "Accepted",
        "PendingAcceptance" => "Invitation pending",
        "" => "-",
        _ => InvitationState,
    };
}
