namespace M365Manager.ViewModels;

/// <summary>
/// A single entry in the left navigation rail.
///
/// <see cref="Section"/> drives the sidebar's grouping (see MainViewModel.PagesView): the rail is
/// ordered by how often the work actually happens, so the six daily pages sit together at the top
/// and the rarely used ones drop below. Changing a Section here changes where the page appears.
/// </summary>
public sealed class NavItem
{
    public required string Title { get; init; }

    /// <summary>Group caption this item appears under, e.g. "DAILY OPERATIONS".</summary>
    public required string Section { get; init; }

    /// <summary>Segoe MDL2 Assets glyph code, e.g. "".</summary>
    public required string Glyph { get; init; }

    public required object ViewModel { get; init; }

    /// <summary>
    /// The accessible name of the rail's list item. Without this a screen reader - and any UI
    /// automation - reads "M365Manager.ViewModels.NavItem" instead of "Groups", because a
    /// templated ListBoxItem falls back to ToString() on its data item.
    /// </summary>
    public override string ToString() => Title;
}
