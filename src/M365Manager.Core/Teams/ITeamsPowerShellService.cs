namespace M365Manager.Core.Teams;

/// <summary>
/// Teams admin operations that only the MicrosoftTeams PowerShell module can do - there is no Graph
/// equivalent for granting a calling policy or assigning a phone number.
///
/// Registered as an <see cref="M365Manager.Core.M365.IM365Connector"/>, so it signs in at app start
/// alongside Exchange and SharePoint.
/// </summary>
public interface ITeamsPowerShellService
{
    bool IsConnected { get; }

    /// <summary>
    /// Assigns a phone number to a user and enables them for enterprise voice in one step
    /// (Set-CsPhoneNumberAssignment). Replaces the deprecated
    /// <c>Set-CsUser -EnterpriseVoiceEnabled</c> the legacy script used.
    /// </summary>
    /// <param name="phoneNumber">E.164 with the leading "+".</param>
    /// <param name="phoneNumberType">"DirectRouting", "CallingPlan" or "OperatorConnect".</param>
    Task AssignPhoneNumberAsync(string upn, string phoneNumber, string phoneNumberType, CancellationToken ct = default);

    /// <summary>
    /// Grants one policy to a user, e.g. Grant-CsOnlineVoiceRoutingPolicy. <paramref name="cmdlet"/>
    /// is the full cmdlet name so the caller keeps the list of policies in one readable place.
    /// </summary>
    Task GrantPolicyAsync(string cmdlet, string upn, string policyName, CancellationToken ct = default);

    /// <summary>
    /// Reads back the user's voice configuration after the changes, so the wizard's result page shows
    /// what Teams actually holds rather than just what was sent to it.
    /// </summary>
    Task<TeamsVoiceState> GetVoiceStateAsync(string upn, CancellationToken ct = default);
}

/// <summary>What Teams reports for a user once the wizard is done. Empty strings mean "not set".</summary>
public sealed class TeamsVoiceState
{
    public string LineUri { get; init; } = "";
    public bool EnterpriseVoiceEnabled { get; init; }
    public string OnlineVoiceRoutingPolicy { get; init; } = "";
    public string TenantDialPlan { get; init; } = "";
    public string TeamsEmergencyCallingPolicy { get; init; } = "";
    public string TeamsEmergencyCallRoutingPolicy { get; init; } = "";
    public string TeamsUpgradeEffectiveMode { get; init; } = "";
}
