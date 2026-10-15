using System.Text.Json;
using M365Manager.Core.M365;

namespace M365Manager.Core.Telephony;

/// <summary>
/// An account's licensing as Entra reports it, read live from <c>/users/{upn}/licenseDetails</c>
/// rather than from the telephony database's nightly <c>e3licensed</c> import - that table only
/// ever knew about E3, which this tenant has since replaced with E5.
///
/// The decisive fact is not which SKU someone holds but whether it grants the <b>MCOEV</b> service
/// plan (Microsoft 365 Phone System / Teams Phone). E5, the Teams Phone Standard add-on, Teams Phone
/// with Calling Plan and Teams Rooms Pro all grant it; E3 never did. Checking the service plan
/// instead of the SKU means this keeps working through the next licensing rename.
/// </summary>
public sealed class TeamsLicenseState
{
    /// <summary>False when the Graph call itself failed - then nothing below means anything.</summary>
    public bool Checked { get; init; }

    /// <summary>Why the lookup failed, when <see cref="Checked"/> is false.</summary>
    public string? Error { get; init; }

    /// <summary>MCOEV is assigned and fully provisioned - the account can be given a number.</summary>
    public bool HasTeamsPhone { get; init; }

    /// <summary>
    /// MCOEV is assigned but Microsoft has not finished provisioning it. Worth flagging, but not
    /// worth blocking on: it usually completes within minutes.
    /// </summary>
    public bool TeamsPhonePending { get; init; }

    /// <summary>The SKUs found, e.g. "SPE_E5" - shown so the operator can see what they actually have.</summary>
    public IReadOnlyList<string> Skus { get; init; } = Array.Empty<string>();

    /// <summary>True when Entra answered and the account has no Teams Phone at all, not even pending.</summary>
    public bool DefinitelyUnlicensed => Checked && !HasTeamsPhone && !TeamsPhonePending;

    /// <summary>How the licensing reads on a lookup card.</summary>
    public string Summary =>
        !Checked ? $"Licence check failed: {Error}"
        : HasTeamsPhone ? $"Teams Phone licensed{SkuSuffix}"
        : TeamsPhonePending ? $"Teams Phone licence still provisioning{SkuSuffix}"
        : $"No Teams Phone licence{SkuSuffix}";

    private string SkuSuffix => Skus.Count == 0 ? " (no licences assigned)" : $" ({string.Join(", ", Skus)})";
}

/// <summary>Reads <see cref="TeamsLicenseState"/> from Graph. Shared by New Hire and room telephony.</summary>
public static class TeamsLicenseReader
{
    /// <summary>
    /// The service plan that actually decides whether an account can hold a phone number. Checking
    /// it rather than the SKU keeps this correct across licensing renames.
    /// </summary>
    public const string TeamsPhoneServicePlan = "MCOEV";

    /// <summary>
    /// A failure here is deliberately not fatal: an unreachable Graph should not stop a run that
    /// would otherwise succeed, and a genuinely missing licence still surfaces when Teams refuses
    /// the number.
    /// </summary>
    public static async Task<TeamsLicenseState> ReadAsync(GraphRestClient graph, string upn, CancellationToken ct)
    {
        try
        {
            var licenses = await graph.GetAllPagesAsync(
                $"/users/{Uri.EscapeDataString(upn)}/licenseDetails", ct: ct);

            var skus = new List<string>();
            var provisioned = false;
            var pending = false;

            foreach (var license in licenses)
            {
                var sku = Text(license, "skuPartNumber");
                if (sku.Length > 0)
                    skus.Add(sku);

                if (!license.TryGetProperty("servicePlans", out var plans) || plans.ValueKind != JsonValueKind.Array)
                    continue;

                foreach (var plan in plans.EnumerateArray())
                {
                    if (!string.Equals(Text(plan, "servicePlanName"), TeamsPhoneServicePlan, StringComparison.OrdinalIgnoreCase))
                        continue;

                    // The same plan can appear under several SKUs; one provisioned copy is enough.
                    if (string.Equals(Text(plan, "provisioningStatus"), "Success", StringComparison.OrdinalIgnoreCase))
                        provisioned = true;
                    else
                        pending = true;
                }
            }

            return new TeamsLicenseState
            {
                Checked = true,
                HasTeamsPhone = provisioned,
                TeamsPhonePending = pending && !provisioned,
                Skus = skus,
            };
        }
        catch (Exception ex)
        {
            return new TeamsLicenseState { Checked = false, Error = ex.Message };
        }
    }

    internal static string Text(JsonElement element, string name) =>
        element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : "";
}
