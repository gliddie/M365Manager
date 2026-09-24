namespace M365Manager.Core.TeamsPolicies;

/// <summary>An Entra group that carries a Teams policy assignment.</summary>
public sealed record TeamsPolicyGroup(string Id, string DisplayName, string Description, string Prefix)
{
    /// <summary>The part of the name after the prefix - what actually distinguishes the policies.</summary>
    public string ShortName =>
        DisplayName.StartsWith(Prefix, StringComparison.OrdinalIgnoreCase)
            ? DisplayName[Prefix.Length..]
            : DisplayName;
}

/// <summary>The user a policy group is being looked up or changed for.</summary>
public sealed record PolicyUser(string Id, string DisplayName, string Upn, string Mail);

/// <summary>What a user currently holds within one configured prefix.</summary>
public sealed class PolicyCategory
{
    public required string Prefix { get; init; }

    /// <summary>Every group with this prefix that exists in the tenant, for the picker.</summary>
    public IReadOnlyList<TeamsPolicyGroup> AvailableGroups { get; init; } = Array.Empty<TeamsPolicyGroup>();

    /// <summary>
    /// What the user is a member of within this prefix. Normally zero or one - more than one means
    /// the tenant drifted from the one-per-prefix rule and the UI says so rather than hiding it.
    /// </summary>
    public IReadOnlyList<TeamsPolicyGroup> CurrentGroups { get; init; } = Array.Empty<TeamsPolicyGroup>();
}

/// <summary>Result of looking a user up.</summary>
public sealed class UserPolicyLookup
{
    public required bool Succeeded { get; init; }
    public PolicyUser? User { get; init; }
    public IReadOnlyList<PolicyCategory> Categories { get; init; } = Array.Empty<PolicyCategory>();
    public string? ErrorMessage { get; init; }

    public static UserPolicyLookup Failed(string error) => new() { Succeeded = false, ErrorMessage = error };
}

/// <summary>
/// Assigns one policy group. Within the group's own prefix this replaces whatever the user had -
/// other prefixes are left alone.
/// </summary>
public sealed class AssignPolicyRequest
{
    public required string TaskNumber { get; init; }
    public required PolicyUser User { get; init; }

    /// <summary>The group to assign. Null removes the user from every group in <see cref="Prefix"/> without adding one.</summary>
    public TeamsPolicyGroup? Group { get; init; }

    /// <summary>Which category is being changed - decides what may be removed.</summary>
    public required string Prefix { get; init; }

    /// <summary>Groups in this prefix the user currently holds; removed unless they are the target.</summary>
    public IReadOnlyList<TeamsPolicyGroup> Replacing { get; init; } = Array.Empty<TeamsPolicyGroup>();
}

public sealed class AssignPolicyResult
{
    public required bool Succeeded { get; init; }
    public IReadOnlyList<string> Added { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> Removed { get; init; } = Array.Empty<string>();
    public string? ErrorMessage { get; init; }

    /// <summary>Set when the assignment went through but a removal did not - the user then holds two groups.</summary>
    public string? WarningMessage { get; init; }

    public Guid CorrelationId { get; init; }

    public static AssignPolicyResult Failed(Guid correlationId, string error) =>
        new() { Succeeded = false, ErrorMessage = error, CorrelationId = correlationId };
}
