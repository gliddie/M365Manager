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

    /// <summary>What Entra says the employee is licensed for.</summary>
    public TeamsLicenseState License { get; init; } = new();

    /// <summary>Why this employee cannot be enabled right now, or null when the run may continue.</summary>
    public string? Blocker { get; init; }

    /// <summary>Notes worth showing that do not stop the run (e.g. no Teams Phone licence recorded).</summary>
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();

    public static NewHireLookup Failed(string error) => new() { ErrorMessage = error };
}

/// <summary>
/// The employee's licensing as Entra reports it, read live from
/// <c>/users/{upn}/licenseDetails</c> rather than from the telephony database's nightly
/// <c>e3licensed</c> import - that table only ever knew about E3, which this tenant has since
/// replaced with E5.
///
/// The decisive fact is not which SKU someone holds but whether it grants the <b>MCOEV</b> service
/// plan (Microsoft 365 Phone System / Teams Phone). E5, the Teams Phone Standard add-on and Teams
/// Phone with Calling Plan all grant it; E3 never did. Checking the service plan instead of the SKU
/// means this keeps working through the next licensing rename.
/// </summary>
public sealed class TeamsLicenseState
{
    /// <summary>False when the Graph call itself failed - then nothing below means anything.</summary>
    public bool Checked { get; init; }

    /// <summary>Why the lookup failed, when <see cref="Checked"/> is false.</summary>
    public string? Error { get; init; }

    /// <summary>MCOEV is assigned and fully provisioned - the user can be given a number.</summary>
    public bool HasTeamsPhone { get; init; }

    /// <summary>
    /// MCOEV is assigned but Microsoft has not finished provisioning it. Worth flagging, but not
    /// worth blocking on: it usually completes within minutes.
    /// </summary>
    public bool TeamsPhonePending { get; init; }

    /// <summary>The SKUs found, e.g. "SPE_E5" - shown so the operator can see what they actually have.</summary>
    public IReadOnlyList<string> Skus { get; init; } = Array.Empty<string>();

    /// <summary>How the licensing reads on the lookup card.</summary>
    public string Summary =>
        !Checked ? $"Licence check failed: {Error}"
        : HasTeamsPhone ? $"Teams Phone licensed{SkuSuffix}"
        : TeamsPhonePending ? $"Teams Phone licence still provisioning{SkuSuffix}"
        : $"No Teams Phone licence{SkuSuffix}";

    private string SkuSuffix => Skus.Count == 0 ? " (no licences assigned)" : $" ({string.Join(", ", Skus)})";
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
