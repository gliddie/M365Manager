using System.Management.Automation;
using M365Manager.Core.M365;
using M365Manager.Core.PowerShell;

namespace M365Manager.Core.Teams;

/// <summary>
/// See <see cref="ITeamsPowerShellService"/>. Same bundled-module + embedded-runspace pattern as
/// <see cref="Exchange.ExchangeService"/> and <see cref="SharePoint.SharePointService"/>, so its
/// commands land in the shared console transcript like everything else.
/// </summary>
public sealed class TeamsPowerShellService : ITeamsPowerShellService, IM365Connector
{
    private readonly PowerShellHost _host;
    private readonly IM365AuthService _auth;

    public TeamsPowerShellService(PowerShellHost host, IM365AuthService auth)
    {
        _host = host;
        _auth = auth;
    }

    public string DisplayName => "Microsoft Teams";

    /// <summary>Last of the connectors - nothing else waits on it, and it is the slowest to sign in.</summary>
    public int Order => 3;

    public bool IsConnected { get; private set; }

    public async Task ConnectAsync(Action<string>? onPrompt = null, CancellationToken ct = default)
    {
        await _host.EnsureTeamsModuleAsync(ct);

        var upn = _auth.CurrentUser?.Upn;

        // Device-code auth, as for Exchange: delegated (so changes stay under the admin's own
        // identity) and workable from a GUI process with no window handle to hand to WAM.
        await _host.InvokeAsync(ps =>
        {
            ps.AddCommand("Connect-MicrosoftTeams")
              .AddParameter("UseDeviceAuthentication", true)
              .AddParameter("ErrorAction", "Stop");
            if (!string.IsNullOrWhiteSpace(upn))
                ps.AddParameter("AccountId", upn);
            // Kept out of the console transcript - the device-code flow prints a one-time code.
        }, onPrompt, ct, suppressTranscript: true);

        IsConnected = true;
    }

    public async Task AssignPhoneNumberAsync(string upn, string phoneNumber, string phoneNumberType, CancellationToken ct = default)
    {
        EnsureConnected();

        await _host.InvokeAsync(ps => ps
            .AddCommand("Set-CsPhoneNumberAssignment")
            .AddParameter("Identity", upn)
            .AddParameter("PhoneNumber", phoneNumber)
            .AddParameter("PhoneNumberType", phoneNumberType)
            .AddParameter("ErrorAction", "Stop"), ct: ct);
    }

    public async Task GrantPolicyAsync(string cmdlet, string upn, string policyName, CancellationToken ct = default)
    {
        EnsureConnected();

        await _host.InvokeAsync(ps => ps
            .AddCommand(cmdlet)
            .AddParameter("Identity", upn)
            .AddParameter("PolicyName", policyName)
            .AddParameter("ErrorAction", "Stop"), ct: ct);
    }

    public async Task<TeamsVoiceState> GetVoiceStateAsync(string upn, CancellationToken ct = default)
    {
        EnsureConnected();

        var results = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-CsOnlineUser")
            .AddParameter("Identity", upn)
            .AddParameter("ErrorAction", "Stop"), ct: ct);

        var user = results.FirstOrDefault();
        if (user is null)
            return new TeamsVoiceState();

        return new TeamsVoiceState
        {
            LineUri = Text(user, "LineUri"),
            EnterpriseVoiceEnabled = string.Equals(Text(user, "EnterpriseVoiceEnabled"), "True", StringComparison.OrdinalIgnoreCase),
            OnlineVoiceRoutingPolicy = PolicyName(user, "OnlineVoiceRoutingPolicy"),
            TenantDialPlan = PolicyName(user, "TenantDialPlan"),
            TeamsEmergencyCallingPolicy = PolicyName(user, "TeamsEmergencyCallingPolicy"),
            TeamsEmergencyCallRoutingPolicy = PolicyName(user, "TeamsEmergencyCallRoutingPolicy"),
            TeamsUpgradeEffectiveMode = Text(user, "TeamsUpgradeEffectiveMode"),
        };
    }

    private void EnsureConnected()
    {
        if (!IsConnected)
            throw new InvalidOperationException(
                "Not connected to Microsoft Teams. Restart the app to sign in, or check the connection status on the dashboard.");
    }

    private static string Text(PSObject source, string property) =>
        source.Properties[property]?.Value?.ToString() ?? "";

    /// <summary>
    /// A granted policy comes back either as a bare string or as an object whose ToString() is
    /// "Tag:&lt;name&gt;" - the tag prefix is noise on a summary page, so it is stripped. An empty
    /// value means the user is on the tenant's global policy.
    /// </summary>
    private static string PolicyName(PSObject source, string property)
    {
        var value = Text(source, property);
        return value.StartsWith("Tag:", StringComparison.OrdinalIgnoreCase) ? value[4..] : value;
    }
}
