using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

namespace M365Manager.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    [ObservableProperty]
    private object? _currentViewModel;

    [ObservableProperty]
    private NavItem? _selectedPage;

    public ObservableCollection<NavItem> Pages { get; }

    public MainViewModel(
        DashboardViewModel dashboard,
        GroupsViewModel groups,
        SettingsViewModel settings,
        LogsViewModel logs)
    {
        Pages = new ObservableCollection<NavItem>
        {
            new() { Title = "Dashboard", ViewModel = dashboard },
            new() { Title = "Groups",    ViewModel = groups },
            new() { Title = "Logs",      ViewModel = logs },
            new() { Title = "Settings",  ViewModel = settings },
        };

        SelectedPage = Pages[0];
    }

    partial void OnSelectedPageChanged(NavItem? value)
    {
        CurrentViewModel = value?.ViewModel;
    }
}
