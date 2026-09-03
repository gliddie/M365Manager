namespace M365Manager.Core.TeamsPolicies;

/// <summary>
/// Teams policy group membership: which of the configured policy groups a user holds, and changing
/// that. The groups themselves are ordinary Entra groups whose display name starts with one of the
/// prefixes in Settings; Teams applies the policy to whoever is a member.
/// </summary>
public interface ITeamsPolicyService
{
    /// <summary>Every configured policy group in the tenant, grouped by prefix. Used to fill the pickers.</summary>
    Task<IReadOnlyList<PolicyCategory>> GetCategoriesAsync(CancellationToken ct = default);

    /// <summary>
    /// Resolves the identity (SamAccountName, UPN or e-mail) and reports which policy group they
    /// hold per prefix.
    /// </summary>
    Task<UserPolicyLookup> LookupUserAsync(string identity, CancellationToken ct = default);

    /// <summary>
    /// Applies one category's change: adds the requested group and removes the other groups the
    /// user held under that same prefix. Never touches groups under a different prefix.
    /// </summary>
    Task<AssignPolicyResult> AssignAsync(AssignPolicyRequest request, Action<string>? onProgress = null, CancellationToken ct = default);
}
