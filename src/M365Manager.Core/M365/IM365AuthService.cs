namespace M365Manager.Core.M365;

public interface IM365AuthService
{
    /// <summary>The currently signed-in admin, or null if not signed in.</summary>
    SignedInUser? CurrentUser { get; }

    /// <summary>True once an interactive sign-in has succeeded.</summary>
    bool IsSignedIn { get; }

    /// <summary>
    /// Signs in interactively with the colleague's own admin account (delegated auth),
    /// using the configured Tenant/Client IDs. Returns the signed-in user.
    /// </summary>
    Task<SignedInUser> SignInAsync(CancellationToken ct = default);

    /// <summary>
    /// Gets an access token for another resource (e.g. Exchange Online) using the same
    /// signed-in credential, so the token carries the colleague's identity.
    /// </summary>
    Task<string> GetAccessTokenAsync(string scope, CancellationToken ct = default);

    /// <summary>
    /// Gets a token that satisfies a claims challenge - an MFA or authentication-context step-up
    /// demanded by a PIM activation policy. Skips the token cache and may show the sign-in window.
    /// </summary>
    Task<string> GetAccessTokenWithClaimsAsync(string scope, string? claims, CancellationToken ct = default);

    /// <summary>
    /// Makes the next token request for every resource fetch a fresh token instead of a cached one.
    /// Needed after a PIM activation: cached tokens still carry the roles from before it.
    /// </summary>
    void InvalidateCachedTokens();

    void SignOut();
}
