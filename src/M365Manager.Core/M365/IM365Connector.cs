namespace M365Manager.Core.M365;

/// <summary>
/// Implemented by every service that needs a signed-in connection to an M365 workload -
/// Graph, Exchange Online, and any future one (Teams, SharePoint, ...). Registering a service
/// under this interface (see CoreServiceCollectionExtensions.AddM365ManagerCore) is enough to
/// have it connected automatically at app startup by <see cref="IStartupConnectionService"/> -
/// no startup or UI code needs to change when a new connector is added.
/// </summary>
public interface IM365Connector
{
    /// <summary>Shown in the startup connection overlay, e.g. "Microsoft Graph", "Exchange Online".</summary>
    string DisplayName { get; }

    /// <summary>
    /// Connectors run in ascending order. Exchange reuses Graph's signed-in UPN as a sign-in
    /// hint, so Graph must run first - give a new connector a value greater than the ones it
    /// depends on.
    /// </summary>
    int Order { get; }

    bool IsConnected { get; }

    /// <summary>
    /// Connects using the signed-in admin's own identity. <paramref name="onPrompt"/> receives
    /// interactive sign-in instructions (device code, browser URL, ...) to surface to the user.
    /// </summary>
    Task ConnectAsync(Action<string>? onPrompt = null, CancellationToken ct = default);
}
