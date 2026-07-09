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

    /// <summary>Lists the members of a group (handles distribution/security and M365 groups).</summary>
    Task<IReadOnlyList<GroupMemberInfo>> GetMembersAsync(DistributionGroupInfo group, CancellationToken ct = default);

    /// <summary>Adds a member (email/UPN) to a group.</summary>
    Task AddMemberAsync(DistributionGroupInfo group, string memberIdentity, CancellationToken ct = default);

    /// <summary>Removes a member from a group.</summary>
    Task RemoveMemberAsync(DistributionGroupInfo group, string memberIdentity, CancellationToken ct = default);
}
