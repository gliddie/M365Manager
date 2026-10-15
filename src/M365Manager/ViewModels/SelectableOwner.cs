using CommunityToolkit.Mvvm.ComponentModel;
using M365Manager.Core.Exchange;

namespace M365Manager.ViewModels;

/// <summary>
/// A current owner with a checkbox - the "which owners does Replace take out" list on the Groups
/// and Shared Mailboxes pages (team request 2026-10-15: it was not visible that Replace swapped out
/// every owner).
/// </summary>
public sealed partial class SelectableOwner : ObservableObject
{
    public SelectableOwner(OwnerEntry owner)
    {
        Owner = owner;
    }

    public OwnerEntry Owner { get; }

    public string Label => Owner.Label;

    [ObservableProperty] private bool _isSelected;
}
