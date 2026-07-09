using Microsoft.Graph;

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

    /// <summary>The authenticated Graph client. Throws if not signed in.</summary>
    GraphServiceClient Graph { get; }

    void SignOut();
}
