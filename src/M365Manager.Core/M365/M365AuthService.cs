using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;
using M365Manager.Core.Settings;
using Microsoft.Identity.Client;

namespace M365Manager.Core.M365;

/// <summary>
/// Delegated (interactive) M365 sign-in. Each colleague authenticates with their OWN
/// admin account, so all Graph calls run under their identity and M365 audit logs
/// attribute changes to them. Works with MFA / Conditional Access.
/// The same credential is reused to get tokens for other services (e.g. Exchange Online).
/// </summary>
public sealed class M365AuthService : IM365AuthService, IM365Connector
{
    public string DisplayName => "Microsoft Graph";
    public int Order => 0;
    bool IM365Connector.IsConnected => IsSignedIn;

    // User.Read.All: resolve another user's directory id when removing a group member.
    // GroupMember.ReadWrite.All: add/remove members of Microsoft 365 (Team-connected) groups via
    // Graph - required because Add-/Remove-UnifiedGroupLinks (Exchange PowerShell) intermittently
    // fails with "We failed to update the group mailbox" on Team-linked groups.
    // Group.ReadWrite.All: create teams/groups and set owners (TeamsService.CreateGroupAsync).
    // Directory.ReadWrite.All: guest-access directory setting on new groups (TeamsService.SetGuestAccessAsync).
    // Mail.Send: owner notification e-mail after team creation (TeamsService.SendOwnerMailAsync)
    // and after shared mailbox creation/owner change (SharedMailboxService).
    // RoleEligibilitySchedule.Read.Directory, RoleAssignmentSchedule.ReadWrite.Directory,
    // RoleManagementPolicy.Read.Directory: PimService finds the signed-in admin's eligible PIM roles,
    // reads each role's maximum activation time and activates them.
    private static readonly string[] Scopes =
    {
        "User.Read", "User.Read.All", "GroupMember.ReadWrite.All",
        "Group.ReadWrite.All", "Directory.ReadWrite.All", "Mail.Send",
        "RoleEligibilitySchedule.Read.Directory", "RoleAssignmentSchedule.ReadWrite.Directory",
        "RoleManagementPolicy.Read.Directory",
    };

    private static readonly HttpClient Http = new();

    private readonly ISettingsService _settings;
    private readonly object _tokenLock = new();
    private InteractiveBrowserCredential? _credential;

    // Set by InvalidateCachedTokens. A token carries the admin roles that were active when it was
    // issued (the "wids" claim) and is cached for up to an hour, so after a PIM activation every
    // resource has to be asked once for a fresh one - otherwise Graph keeps refusing for that hour.
    private DateTimeOffset? _tokensInvalidatedAt;
    private readonly HashSet<string> _scopesRefreshedSinceInvalidation = new(StringComparer.OrdinalIgnoreCase);

    public M365AuthService(ISettingsService settings)
    {
        _settings = settings;
    }

    public SignedInUser? CurrentUser { get; private set; }

    public bool IsSignedIn => CurrentUser is not null;

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

        // The first token asks for the app's delegated scopes by name, as the Graph SDK client
        // used to - that is the request that shows the sign-in (and any consent) prompt. Every
        // later call goes through GraphRestClient with .default. The SDK itself is gone: /me here
        // was the last thing it was used for, and it is one GET.
        var token = await _credential.GetTokenAsync(
            new TokenRequestContext(Scopes.Select(s => $"https://graph.microsoft.com/{s}").ToArray()), ct);

        using var request = new HttpRequestMessage(HttpMethod.Get,
            "https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName,displayName");
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

        using var response = await Http.SendAsync(request, ct);
        var payload = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode)
        {
            _credential = null;
            throw new InvalidOperationException(
                $"Signed in, but reading the account from Microsoft Graph failed ({(int)response.StatusCode}): "
                + (payload.Length > 300 ? payload[..300] + "..." : payload));
        }

        using var me = JsonDocument.Parse(payload);
        string Text(string name) =>
            me.RootElement.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                ? value.GetString() ?? ""
                : "";

        CurrentUser = new SignedInUser(Text("userPrincipalName"), Text("displayName"), Text("id"));
        return CurrentUser;
    }

    /// <summary>IM365Connector entry point used by the startup connector - just wraps <see cref="SignInAsync"/>.</summary>
    public Task ConnectAsync(Action<string>? onPrompt = null, CancellationToken ct = default)
    {
        onPrompt?.Invoke("Opening a browser window to sign in to Microsoft 365...");
        return SignInAsync(ct);
    }

    public async Task<string> GetAccessTokenAsync(string scope, CancellationToken ct = default)
    {
        string? claims = null;
        lock (_tokenLock)
        {
            if (_tokensInvalidatedAt is { } since && !_scopesRefreshedSinceInvalidation.Contains(scope))
                claims = NotBeforeClaims(since);
        }

        var token = await GetAccessTokenWithClaimsAsync(scope, claims, ct);

        if (claims is not null)
            lock (_tokenLock) _scopesRefreshedSinceInvalidation.Add(scope);

        return token;
    }

    public async Task<string> GetAccessTokenWithClaimsAsync(string scope, string? claims, CancellationToken ct = default)
    {
        if (_credential is null)
            throw new InvalidOperationException("Not signed in to M365. Sign in first.");

        // With claims the token cache is skipped: the credential redeems its refresh token, or -
        // for an MFA or authentication-context challenge - opens the sign-in window again.
        var context = new TokenRequestContext(new[] { scope }, claims: claims);
        var token = await _credential.GetTokenAsync(context, ct);
        return token.Token;
    }

    public async Task<string> GetAccessTokenWithFreshMfaAsync(string scope, CancellationToken ct = default)
    {
        var m365 = _settings.Current.M365;
        if (_credential is null || string.IsNullOrWhiteSpace(m365.ClientId))
            throw new InvalidOperationException("Not signed in to M365. Sign in first.");

        // Straight MSAL, because InteractiveBrowserCredential cannot pass extra query parameters.
        // amr_values=ngcmfa makes Entra ask for MFA now even when the browser session would sign in
        // silently - the same prompt the Entra portal shows before a PIM activation. A claims request
        // for amr=mfa does not do that: Entra answers it from the refresh token, without MFA.
        var app = PublicClientApplicationBuilder.Create(m365.ClientId)
            .WithTenantId(m365.TenantId)
            .WithRedirectUri("http://localhost")
            .Build();

        var request = app.AcquireTokenInteractive(new[] { scope })
            .WithUseEmbeddedWebView(false)
            .WithExtraQueryParameters(new Dictionary<string, (string, bool)> { ["amr_values"] = ("ngcmfa", false) });
        if (CurrentUser?.Upn is { Length: > 0 } upn)
            request = request.WithLoginHint(upn);

        var result = await request.ExecuteAsync(ct);
        return result.AccessToken;
    }

    public void InvalidateCachedTokens()
    {
        lock (_tokenLock)
        {
            _tokensInvalidatedAt = DateTimeOffset.UtcNow;
            _scopesRefreshedSinceInvalidation.Clear();
        }
    }

    /// <summary>
    /// The claims challenge Entra itself sends when it wants a token issued after a point in time
    /// (continuous access evaluation). Asking for it is the supported way to force a fresh token.
    /// </summary>
    private static string NotBeforeClaims(DateTimeOffset since) =>
        "{\"access_token\":{\"nbf\":{\"essential\":true,\"value\":\"" + since.ToUnixTimeSeconds() + "\"}}}";

    /// <summary>
    /// Graph stays signed in on a reconnect - there is no session to tear down, only tokens that
    /// may carry roles which have since expired. They are refreshed on next use instead.
    /// </summary>
    Task IM365Connector.DisconnectAsync(CancellationToken ct)
    {
        InvalidateCachedTokens();
        return Task.CompletedTask;
    }

    public void SignOut()
    {
        _credential = null;
        CurrentUser = null;
    }
}
