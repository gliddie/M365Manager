using Azure.Core;
using Azure.Identity;
using M365Manager.Core.Settings;
using Microsoft.Graph;

namespace M365Manager.Core.M365;

/// <summary>
/// Delegated (interactive) M365 sign-in. Each colleague authenticates with their OWN
/// admin account, so all Graph calls run under their identity and M365 audit logs
/// attribute changes to them. Works with MFA / Conditional Access.
/// The same credential is reused to get tokens for other services (e.g. Exchange Online).
/// </summary>
public sealed class M365AuthService : IM365AuthService
{
    // User.Read.All: resolve another user's directory id when removing a group member.
    // GroupMember.ReadWrite.All: add/remove members of Microsoft 365 (Team-connected) groups via
    // Graph - required because Add-/Remove-UnifiedGroupLinks (Exchange PowerShell) intermittently
    // fails with "We failed to update the group mailbox" on Team-linked groups.
    private static readonly string[] Scopes = { "User.Read", "User.Read.All", "GroupMember.ReadWrite.All" };

    private readonly ISettingsService _settings;
    private GraphServiceClient? _graph;
    private InteractiveBrowserCredential? _credential;

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

        _credential = new InteractiveBrowserCredential(options);
        _graph = new GraphServiceClient(_credential, Scopes);

        var me = await _graph.Me.GetAsync(cancellationToken: ct);

        CurrentUser = new SignedInUser(
            me?.UserPrincipalName ?? "",
            me?.DisplayName ?? "",
            me?.Id ?? "");

        return CurrentUser;
    }

    public async Task<string> GetAccessTokenAsync(string scope, CancellationToken ct = default)
    {
        if (_credential is null)
            throw new InvalidOperationException("Not signed in to M365. Sign in first.");

        var context = new TokenRequestContext(new[] { scope });
        var token = await _credential.GetTokenAsync(context, ct);
        return token.Token;
    }

    public void SignOut()
    {
        _graph = null;
        _credential = null;
        CurrentUser = null;
    }
}
