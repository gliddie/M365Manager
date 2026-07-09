namespace M365Manager.Core.Exchange;

public interface IExchangeService
{
    bool IsConnected { get; }

    /// <summary>
    /// Connects to Exchange Online using an access token from the signed-in admin
    /// (same browser login as Graph) - no device code, actions run under their identity.
    /// </summary>
    Task ConnectAsync(CancellationToken ct = default);

    /// <summary>Finds distribution, mail-enabled security and Microsoft 365 groups matching the text.</summary>
    Task<IReadOnlyList<DistributionGroupInfo>> SearchGroupsAsync(string search, CancellationToken ct = default);

    /// <summary>Lists the members of a group (handles distribution/security and M365 groups).</summary>
    Task<IReadOnlyList<GroupMemberInfo>> GetMembersAsync(DistributionGroupInfo group, CancellationToken ct = default);
}
