using M365Manager.Core.ActiveDirectory;
using M365Manager.Core.Teams;
using M365Manager.Data.NewHire;

namespace M365Manager.Core.NewHire;

/// <summary>
/// What the wizard found out about an employee in step 1: their AD account, their licence state, and
/// the site that decides which numbers and policies they get.
///
/// <see cref="Blocker"/> carries the one reason the run cannot continue, if there is one. The legacy
/// tool showed each of these as a separate message box and then let the operator carry on regardless.
/// </summary>
public sealed class NewHireLookup
{
    public bool Succeeded { get; init; }
    public string? ErrorMessage { get; init; }

    public AdUser? User { get; init; }

    /// <summary>The site matched from the user's AD office, or the one the operator picked instead.</summary>
    public UcLocation? Location { get; init; }

    public UcEmployeeState State { get; init; } = new();

    /// <summary>Why this employee cannot be enabled right now, or null when the run may continue.</summary>
    public string? Blocker { get; init; }

    /// <summary>Notes worth showing that do not stop the run (e.g. no Teams Phone licence recorded).</summary>
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();

    public static NewHireLookup Failed(string error) => new() { ErrorMessage = error };
}

/// <summary>Everything the operator has assembled by the end of step 3, ready to be applied.</summary>
public sealed class NewHireRequest
{
    /// <summary>ServiceNow ticket authorizing the change. Required.</summary>
    public required string TaskNumber { get; init; }

    public required AdUser User { get; init; }

    public required UcLocation Location { get; init; }

    /// <summary>The number chosen in step 2, bare digits without "+".</summary>
    public required string Did { get; init; }

    /// <summary>tel:+&lt;did&gt; - what goes into msRTCSIP-Line.</summary>
    public string LineUri => $"tel:+{Did}";

    /// <summary>E.164 - what Set-CsPhoneNumberAssignment wants.</summary>
    public string PhoneNumber => $"+{Did}";

    /// <summary>sip:&lt;mail&gt; - what goes into msRTCSIP-PrimaryUserAddress, as the legacy script derived it.</summary>
    public string SipAddress => $"sip:{User.Mail}";

    /// <summary>
    /// All four voice policies are named after the site code - that is the convention in the tenant
    /// and how the legacy script granted them.
    /// </summary>
    public string PolicyName => Location.LocationCode;
}

/// <summary>One thing the wizard did, with its outcome. Drives the result page and the run log.</summary>
/// <param name="Name">What was attempted, e.g. "Voice routing policy".</param>
/// <param name="Detail">The value involved, e.g. the policy name or the number.</param>
/// <param name="Error">Null when the step succeeded.</param>
public sealed record NewHireStep(string Name, string Detail, string? Error)
{
    public bool Succeeded => Error is null;

    public static NewHireStep Ok(string name, string detail) => new(name, detail, null);
    public static NewHireStep Failed(string name, string detail, string error) => new(name, detail, error);
}

/// <summary>Outcome of a whole run.</summary>
public sealed class NewHireResult
{
    public bool Succeeded { get; init; }
    public string? ErrorMessage { get; init; }

    /// <summary>Every step attempted, in order.</summary>
    public IReadOnlyList<NewHireStep> Steps { get; init; } = Array.Empty<NewHireStep>();

    /// <summary>What Teams reports for the user afterwards. Null when the read-back itself failed.</summary>
    public TeamsVoiceState? VoiceState { get; init; }

    /// <summary>Row id in the telephony database's newhirelog table, for cross-referencing.</summary>
    public int NewHireLogId { get; init; }

    /// <summary>Ties every audit log line of this run together.</summary>
    public Guid CorrelationId { get; init; }

    /// <summary>True when the employee was told about their new number.</summary>
    public bool EmployeeNotified { get; init; }

    /// <summary>
    /// Teams refused the number because it is managed on-premises, but the AD write succeeded - so
    /// the number reaches Teams on the next Entra Connect cycle rather than immediately. The run is
    /// fine; the result page just has to say why Teams does not show it yet.
    /// </summary>
    public bool NumberPendingSync { get; init; }

    public bool HasFailedSteps => Steps.Any(s => !s.Succeeded);

    public static NewHireResult Failed(Guid correlationId, string error) =>
        new() { CorrelationId = correlationId, ErrorMessage = error };
}
