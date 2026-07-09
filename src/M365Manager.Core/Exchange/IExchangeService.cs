namespace M365Manager.Core.Exchange;

public interface IExchangeService
{
    bool IsConnected { get; }

    /// <summary>
    /// Connects to Exchange Online as the signed-in admin using device-code auth.
    /// <paramref name="onPrompt"/> receives the sign-in instructions (URL + code) to show the user.
    /// </summary>
    Task ConnectAsync(Action<string>? onPrompt = null, CancellationToken ct = default);

    /// <summary>Finds distribution/security groups matching the search text.</summary>
    Task<IReadOnlyList<DistributionGroupInfo>> SearchGroupsAsync(string search, CancellationToken ct = default);

    /// <summary>Gets the full details of one group by name/alias/address.</summary>
    Task<DistributionGroupInfo?> GetGroupAsync(string identity, CancellationToken ct = default);

    /// <summary>Lists the members of a group.</summary>
    Task<IReadOnlyList<GroupMemberInfo>> GetMembersAsync(string identity, CancellationToken ct = default);
}
