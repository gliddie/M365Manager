namespace M365Manager.ViewModels;

/// <summary>A single entry in the left navigation rail.</summary>
public sealed class NavItem
{
    public required string Title { get; init; }

    public required object ViewModel { get; init; }
}
