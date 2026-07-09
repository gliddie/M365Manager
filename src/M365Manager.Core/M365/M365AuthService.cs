using Azure.Identity;
using M365Manager.Core.Settings;
using Microsoft.Graph;

namespace M365Manager.Core.M365;

/// <summary>
/// Delegated (interactive) M365 sign-in. Each colleague authenticates with their OWN
/// admin account, so all Graph calls run under their identity and M365 audit logs
/// attribute changes to them. Works with MFA / Conditional Access.
/// </summary>
public sealed class M365AuthService : IM365AuthService
{
    private static readonly string[] Scopes = { "User.Read" };

    private readonly ISettingsService _settings;
    private GraphServiceClient? _graph;

    public M365AuthService(ISettingsService settings)
    {
        _settings = settings;
    }

    public SignedInUser? CurrentUser { get; private set; }

    public bool IsSignedIn => CurrentUser is not null;

    public GraphServiceClient Graph =>
        _graph ?? throw new InvalidOperationException("Not signed in to M365. Sign in first.");

    public async Task<SignedInUser> SignInAsync(CancellationToken ct = default)
    {
        var m365 = _settings.Current.M365;

        if (string.IsNullOrWhiteSpace(m365.TenantId) || string.IsNullOrWhiteSpace(m365.ClientId))
            throw new InvalidOperationException(
                "M365 is not configured. Enter the Tenant ID and Application (Client) ID in Settings.");

        var options = new InteractiveBrowserCredentialOptions
        {
            TenantId = m365.TenantId,
            ClientId = m365.ClientId,
            RedirectUri = new Uri("http://localhost"),
        };

        var credential = new InteractiveBrowserCredential(options);
        _graph = new GraphServiceClient(credential, Scopes);

        var me = await _graph.Me.GetAsync(cancellationToken: ct);

        CurrentUser = new SignedInUser(
            me?.UserPrincipalName ?? "",
            me?.DisplayName ?? "",
            me?.Id ?? "");

        return CurrentUser;
    }

    public void SignOut()
    {
        _graph = null;
        CurrentUser = null;
    }
}
