using M365Manager.Core.Teams;
using M365Manager.Core.Telephony;
using M365Manager.Data.NewHire;

namespace M365Manager.Core.RoomResources;

/// <summary>
/// What the telephony lookup found out about a room account: its Entra identity, licence, usage
/// location and current voice state, plus the site the room's code points at. <see cref="Blocker"/>
/// carries the one reason it cannot be given a number right now.
/// </summary>
public sealed class RoomTelephonyLookup
{
    public bool Succeeded { get; init; }
    public string? ErrorMessage { get; init; }

    public string UserPrincipalName { get; init; } = "";
    public string DisplayName { get; init; } = "";

    /// <summary>The usage location Entra holds now (ISO code), or "" when none is set - the default for a new room.</summary>
    public string CurrentUsageLocation { get; init; } = "";

    /// <summary>New-Mailbox -Room creates the account disabled; a Teams Rooms device cannot sign in until it is enabled.</summary>
    public bool AccountEnabled { get; init; }

    public TeamsLicenseState License { get; init; } = new();

    /// <summary>What Teams reports for the account; null when Teams does not know it (yet).</summary>
    public TeamsVoiceState? Voice { get; init; }

    /// <summary>The number the room already holds - in Teams or in the telephony inventory - or null.</summary>
    public string? AssignedNumber { get; init; }

    /// <summary>The uchelper site whose location code equals the room's site code, when there is one.</summary>
    public UcLocation? SuggestedLocation { get; init; }

    /// <summary>
    /// Set when the room has no Teams Phone licence but the tenant has a free Teams Rooms licence
    /// that carries it - the tab then offers to assign it, after the operator confirms.
    /// </summary>
    public RoomLicenceOffer? LicenceOffer { get; init; }

    public string? Blocker { get; init; }
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();

    public static RoomTelephonyLookup Failed(string error) => new() { ErrorMessage = error };
}

/// <summary>A Teams Rooms licence the tenant holds that includes Teams Phone, with how many are free.</summary>
public sealed record RoomLicenceOffer(Guid SkuId, string SkuPartNumber, int Enabled, int Consumed)
{
    public int Available => Math.Max(0, Enabled - Consumed);

    /// <summary>"Microsoft_Teams_Rooms_Pro" → "Microsoft Teams Rooms Pro".</summary>
    public string DisplayName => SkuPartNumber.Replace('_', ' ');

    public string StockText => $"{Available} of {Enabled} {DisplayName} licence(s) free";
}

/// <summary>Telephony asked for on the Create form, carried on <see cref="RoomResourceCreationRequest.Telephony"/>.</summary>
public sealed class RoomCreateTelephony
{
    public required UcLocation Location { get; init; }
    public required UcDidRange Range { get; init; }
    public required string Did { get; init; }

    public string PhoneNumber => $"+{Did}";
}

/// <summary>
/// A Teams Rooms licence asked for on the Create form - independent of telephony (user request
/// 2026-10-15). Set only after the operator confirmed the dialog.
/// </summary>
public sealed class RoomCreateLicence
{
    /// <summary>Its country becomes the usage location, which Entra needs before any licence.</summary>
    public required UcLocation Location { get; init; }

    public required RoomLicenceOffer Offer { get; init; }
}

/// <summary>The confirmed request to give a room the offered Teams Rooms licence.</summary>
public sealed class RoomLicenceRequest
{
    public required string TaskNumber { get; init; }
    public required string UserPrincipalName { get; init; }
    public required string CurrentUsageLocation { get; init; }

    /// <summary>Its country becomes the usage location, which Entra needs before any licence.</summary>
    public required UcLocation Location { get; init; }

    public required RoomLicenceOffer Offer { get; init; }
}

/// <summary>Everything assembled on the Telephony tab, ready to be applied.</summary>
public sealed class RoomTelephonyRequest
{
    /// <summary>ServiceNow ticket authorizing the change. Required.</summary>
    public required string TaskNumber { get; init; }

    public required string UserPrincipalName { get; init; }
    public required string DisplayName { get; init; }

    /// <summary>The usage location found by the lookup, to tell whether it has to change.</summary>
    public required string CurrentUsageLocation { get; init; }

    /// <summary>The uchelper site: decides the number blocks, the policy names and the usage location.</summary>
    public required UcLocation Location { get; init; }

    /// <summary>The block the number comes from; its type decides Direct Routing or Operator Connect.</summary>
    public required UcDidRange Range { get; init; }

    /// <summary>Bare digits without "+".</summary>
    public required string Did { get; init; }

    public string PhoneNumber => $"+{Did}";
    public string LineUri => $"tel:+{Did}";

    public bool IsOperatorConnect => Range.IsOperatorConnect;

    /// <summary>The voice policies are named after the site's location code - the tenant's convention.</summary>
    public string PolicyName => Location.LocationCode;

    public IReadOnlyList<(string Name, string Cmdlet)> VoicePolicies => PhonePolicyPlan.For(Range.Type);

    /// <summary>The site's country as Microsoft 365 wants it - the room is set to it when it differs.</summary>
    public UsageLocation TargetUsageLocation => UsageLocationResolver.Resolve(Location.Country, Location.Name);

    public bool ChangesUsageLocation =>
        TargetUsageLocation.Code is { } target && !string.Equals(target, CurrentUsageLocation, StringComparison.OrdinalIgnoreCase);
}

/// <summary>One thing the telephony setup did, with its outcome.</summary>
/// <param name="Error">Null when the step succeeded.</param>
public sealed record RoomTelephonyStep(string Name, string Detail, string? Error)
{
    public bool Succeeded => Error is null;

    public static RoomTelephonyStep Ok(string name, string detail) => new(name, detail, null);
    public static RoomTelephonyStep Failed(string name, string detail, string error) => new(name, detail, error);
}

public sealed class RoomTelephonyResult
{
    /// <summary>True when the room holds its number; failed policy steps are reported but do not undo that.</summary>
    public bool Succeeded { get; init; }
    public string? ErrorMessage { get; init; }

    public IReadOnlyList<RoomTelephonyStep> Steps { get; init; } = Array.Empty<RoomTelephonyStep>();

    /// <summary>What Teams reports afterwards. Null when the read-back itself failed.</summary>
    public TeamsVoiceState? VoiceState { get; init; }

    public Guid CorrelationId { get; init; }

    public bool HasFailedSteps => Steps.Any(s => !s.Succeeded);

    public static RoomTelephonyResult Failed(Guid correlationId, string error) =>
        new() { CorrelationId = correlationId, ErrorMessage = error };
}
