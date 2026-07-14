namespace M365Manager.Core.Exchange;

public interface IExchangeService
{
    bool IsConnected { get; }

    /// <summary>
    /// Connects to Exchange Online as the signed-in admin using device-code auth
    /// (delegated - actions run under their identity). <paramref name="onPrompt"/>
    /// receives the sign-in instructions (URL + code) to display to the user.
    /// </summary>
    Task ConnectAsync(Action<string>? onPrompt = null, CancellationToken ct = default);

    /// <summary>Finds distribution, mail-enabled security and Microsoft 365 groups matching the text.</summary>
    Task<IReadOnlyList<DistributionGroupInfo>> SearchGroupsAsync(string search, CancellationToken ct = default);

    /// <summary>Tries to resolve groups from the SQL cache first and falls back to live M365 lookup.</summary>
    Task<IReadOnlyList<DistributionGroupInfo>> SearchGroupsCachedAsync(string search, CancellationToken ct = default);

    /// <summary>Lists the members of a group (handles distribution/security and M365 groups).</summary>
    Task<IReadOnlyList<GroupMemberInfo>> GetMembersAsync(DistributionGroupInfo group, CancellationToken ct = default);

    /// <summary>
    /// Adds one or more members (email/UPN/SamAccountName) to a group. A bare SamAccountName
    /// (no "@") is completed with the configured default domain. Each identity is attempted
    /// independently - one failure doesn't stop the rest - and the cache is synced once at the end.
    /// </summary>
    Task<IReadOnlyList<MemberOperationResult>> AddMembersAsync(DistributionGroupInfo group, IEnumerable<string> memberIdentities, CancellationToken ct = default);

    /// <summary>
    /// Removes one or more members from a group. Each identity is attempted independently - one
    /// failure doesn't stop the rest - and the cache is synced once at the end.
    /// </summary>
    Task<IReadOnlyList<MemberOperationResult>> RemoveMembersAsync(DistributionGroupInfo group, IEnumerable<string> memberIdentities, CancellationToken ct = default);

    /// <summary>
    /// Re-fetches the group's direct members live from Exchange and refreshes the SQL cache
    /// (dbo.GroupMembers) for them, so a change made here or by another tool/admin shows up
    /// immediately instead of waiting for the next scheduled import. Nested-member rows produced
    /// by the import script are left untouched since this only sees direct membership.
    /// </summary>
    Task<IReadOnlyList<GroupMemberInfo>> SyncMembersAsync(DistributionGroupInfo group, CancellationToken ct = default);
}
