using M365Manager.Core.M365;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;

namespace M365Manager.Core.SharePoint;

/// <summary>
/// SharePoint Online admin connector (PnP.PowerShell, embedded runspace). Reuses the same
/// bundled-module + interactive-auth pattern as <see cref="Exchange.ExchangeService"/>.
/// Registered as a low-priority <see cref="IM365Connector"/> (Order = 3) so it auto-connects
/// at startup alongside Graph and Exchange - see CoreServiceCollectionExtensions.
/// </summary>
public sealed class SharePointService : ISharePointService, IM365Connector
{
    private readonly PowerShellHost _host;
    private readonly IM365AuthService _auth;
    private readonly ISettingsService _settings;

    public SharePointService(PowerShellHost host, IM365AuthService auth, ISettingsService settings)
    {
        _host = host;
        _auth = auth;
        _settings = settings;
    }

    public string DisplayName => "SharePoint Online";
    public int Order => 3;

    public bool IsConnected { get; private set; }

    public async Task ConnectAsync(Action<string>? onPrompt = null, CancellationToken ct = default)
    {
        var m365 = _settings.Current.M365;
        if (string.IsNullOrWhiteSpace(m365.SharePointAdminUrl))
            throw new InvalidOperationException(
                "SharePoint admin center URL is not configured. Enter it in Settings (e.g. https://contoso-admin.sharepoint.com).");
        if (string.IsNullOrWhiteSpace(m365.ClientId))
            throw new InvalidOperationException("M365 is not configured. Enter the Application (Client) ID in Settings.");

        await _host.EnsurePnPModuleAsync(ct);

        // Interactive auth via the same app registration used for Graph/Exchange (AllSites.FullControl
        // delegated permission under the "SharePoint" API, not Graph) - keeps a colleague's
        // own identity on every change. A pure
        // access-token passthrough was tried for Exchange in this project and failed ("UnAuthorized"),
        // so this follows Exchange's proven interactive pattern instead of -AccessToken.
        await _host.InvokeAsync(ps => ps
            .AddCommand("Connect-PnPOnline")
            .AddParameter("Url", m365.SharePointAdminUrl)
            .AddParameter("ClientId", m365.ClientId)
            .AddParameter("Interactive", true)
            // Kept out of the console transcript - this is a sign-in.
            .AddParameter("ErrorAction", "Stop"), onPrompt, ct, suppressTranscript: true);

        IsConnected = true;
    }

    public async Task DisconnectAsync(CancellationToken ct = default)
    {
        if (!IsConnected)
            return;

        IsConnected = false;
        try
        {
            await _host.InvokeAsync(ps => ps
                .AddCommand("Disconnect-PnPOnline")
                .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);
        }
        catch
        {
            // Already gone is fine - the reconnect opens a new session either way.
        }
    }

    public async Task LockSiteSharingAsync(string siteUrl, CancellationToken ct = default)
    {
        if (!IsConnected)
            throw new InvalidOperationException("Not connected to SharePoint Online. Sign in first.");

        await _host.InvokeAsync(ps => ps
            .AddCommand("Set-PnPTenantSite")
            .AddParameter("Identity", siteUrl)
            .AddParameter("SharingCapability", "Disabled")
            .AddParameter("ErrorAction", "Stop"), ct: ct);
    }
}
