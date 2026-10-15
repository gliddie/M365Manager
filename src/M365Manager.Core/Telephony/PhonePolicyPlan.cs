using M365Manager.Data.NewHire;

namespace M365Manager.Core.Telephony;

/// <summary>
/// Which voice policies an account gets for its number, by the number's type. Shared by New Hire and
/// the room telephony setup so the review screens and the runs cannot drift apart.
///
/// Ported from Switchboard's PhonePolicyPlan (common area phone / employee rules):
///   - Direct Routing: voice routing policy, dial plan, emergency calling and emergency call routing
///     policy - all named after the site's location code, the tenant's convention.
///   - Operator Connect: the dial plan only. The operator routes the calls, so there is no voice
///     routing policy, and no emergency policies either (confirmed by the user 2026-10-12 for
///     Switchboard; the same rule applies here).
/// </summary>
public static class PhonePolicyPlan
{
    public const string VoiceRoutingCmdlet = "Grant-CsOnlineVoiceRoutingPolicy";
    public const string DialPlanCmdlet = "Grant-CsTenantDialPlan";
    public const string EmergencyCallingCmdlet = "Grant-CsTeamsEmergencyCallingPolicy";
    public const string EmergencyCallRoutingCmdlet = "Grant-CsTeamsEmergencyCallRoutingPolicy";

    private static readonly (string Name, string Cmdlet) DialPlan = ("Dial plan", DialPlanCmdlet);

    private static readonly (string Name, string Cmdlet)[] DirectRouting =
    {
        ("Voice routing policy", VoiceRoutingCmdlet),
        DialPlan,
        ("Emergency calling policy", EmergencyCallingCmdlet),
        ("Emergency call routing policy", EmergencyCallRoutingCmdlet),
    };

    /// <summary>The policies to grant, in order. Empty for a range whose type is not set.</summary>
    public static IReadOnlyList<(string Name, string Cmdlet)> For(TeamsNumberType type) => type switch
    {
        TeamsNumberType.DirectRouting => DirectRouting,
        TeamsNumberType.OperatorConnect => new[] { DialPlan },
        _ => Array.Empty<(string, string)>(),
    };

    /// <summary>True when <paramref name="cmdlet"/> is part of the plan for <paramref name="type"/>.</summary>
    public static bool Grants(TeamsNumberType type, string cmdlet) => For(type).Any(p => p.Cmdlet == cmdlet);

    /// <summary>One line for a review screen, e.g. "Voice routing policy, Dial plan, ... (BRM)".</summary>
    public static string Describe(TeamsNumberType type, string locationCode)
    {
        var list = For(type);
        return list.Count == 0
            ? "None - the number range has no type"
            : $"{string.Join(", ", list.Select(p => p.Name))}: {locationCode}"
              + (type == TeamsNumberType.OperatorConnect ? " (Operator Connect: the operator routes calls, so no voice routing or emergency policies)" : "");
    }
}
