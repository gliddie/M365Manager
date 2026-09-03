using M365Manager.Data.NewHire;

namespace M365Manager.Core.NewHire;

/// <summary>
/// Enables a new employee for Teams telephony: finds their AD account and site, hands out a free
/// number from the site's DID ranges, writes the number to on-premises AD (when the SfB attributes
/// are still synchronised), grants the site's voice policies in Teams, and records the run.
///
/// Ported from the standalone NewHireWizard. The step order is deliberate and is the one thing that
/// must not be rearranged: AD first, then Teams. Where Entra Connect still synchronises the
/// msRTCSIP-* attributes, on-premises is the authoritative source - a number assigned only in Teams
/// gets overwritten, or cleared, on the next sync cycle. Writing both sides with the same value
/// makes the sync a no-op. See <see cref="Settings.NewHireSettings.WriteAdAttributes"/>.
/// </summary>
public interface INewHireService
{
    /// <summary>The sites available in the office picker, ordered by name.</summary>
    Task<IReadOnlyList<UcLocation>> GetLocationsAsync(CancellationToken ct = default);

    /// <summary>
    /// Step 1: resolves an employee and works out whether they can be enabled. Never throws for an
    /// employee-shaped problem (unknown account, no licence, site not enabled for voice) - those come
    /// back as <see cref="NewHireLookup.Blocker"/> so the page can explain them.
    /// </summary>
    /// <param name="officeOverride">
    /// Set when the operator picked a different site than the one on the AD account - the office
    /// attribute is often wrong or empty for a brand-new hire.
    /// </param>
    Task<NewHireLookup> LookUpAsync(string identity, string? officeOverride = null, CancellationToken ct = default);

    /// <summary>Step 2: the site's usable number blocks, least-used first.</summary>
    Task<IReadOnlyList<UcDidRange>> GetDidRangesAsync(string locationCode, CancellationToken ct = default);

    /// <summary>Step 2: the numbers still free inside one block.</summary>
    Task<IReadOnlyList<FreeDid>> FindFreeDidsAsync(UcDidRange range, CancellationToken ct = default);

    /// <summary>
    /// Step 3: applies everything. Individual steps that fail are reported in
    /// <see cref="NewHireResult.Steps"/> rather than aborting the run - a half-configured account
    /// needs the remaining steps attempted, and the operator needs to see exactly which one broke.
    /// </summary>
    Task<NewHireResult> ConfigureAsync(NewHireRequest request, Action<string>? onProgress = null, CancellationToken ct = default);
}
