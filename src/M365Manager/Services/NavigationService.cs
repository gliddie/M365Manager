namespace M365Manager.Services;

/// <summary>One request to move to a page, optionally landing on a particular tab.</summary>
/// <param name="PageTitle">Must match a <see cref="ViewModels.NavItem.Title"/> exactly.</param>
/// <param name="TabIndex">Zero-based tab to select on arrival, or null to leave it as it was.</param>
public readonly record struct NavigationRequest(string PageTitle, int? TabIndex = null);

public interface INavigationService
{
    /// <summary>Raised when something asks to move. The shell listens; nothing else should.</summary>
    event Action<NavigationRequest>? Requested;

    /// <summary>Moves to the named page, optionally selecting a tab once there.</summary>
    void GoTo(string pageTitle, int? tabIndex = null);
}

/// <summary>
/// Lets a page ask the shell to navigate without holding a reference to it.
///
/// A direct reference would not work: MainViewModel is constructed with every page view model as a
/// dependency, so a page that depended on MainViewModel would close the circle. An event keeps the
/// arrow pointing one way - pages raise, the shell listens.
/// </summary>
public sealed class NavigationService : INavigationService
{
    public event Action<NavigationRequest>? Requested;

    public void GoTo(string pageTitle, int? tabIndex = null)
        => Requested?.Invoke(new NavigationRequest(pageTitle, tabIndex));
}
