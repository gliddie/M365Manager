namespace M365Manager.Core.SharePoint;

/// <summary>
/// SharePoint Online tenant-admin operations via the hosted PowerShell runspace
/// (bundled PnP.PowerShell module). Used by the Teams feature to lock down external
/// sharing on the SharePoint site behind an internal team - something a normal team
/// creator can't do themselves (see Set-SPOSite/Set-PnPTenantSite in the legacy script).
/// </summary>
public interface ISharePointService
{
    bool IsConnected { get; }

    /// <summary>
    /// Connects to the SharePoint admin center as the signed-in admin using the same
    /// interactive app-registration flow as Exchange/Graph. <paramref name="onPrompt"/>
    /// receives sign-in instructions to display to the user.
    /// </summary>
    Task ConnectAsync(Action<string>? onPrompt = null, CancellationToken ct = default);

    /// <summary>Disables external sharing on the given site (SharingCapability = Disabled).</summary>
    Task LockSiteSharingAsync(string siteUrl, CancellationToken ct = default);
}
