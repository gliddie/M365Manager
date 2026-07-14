using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Core.M365;
using M365Manager.Services;

namespace M365Manager.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private readonly IStartupConnectionService _startup;
    private readonly DashboardViewModel _dashboard;
    private readonly GroupsViewModel _groups;

    [ObservableProperty]
    private object? _currentViewModel;

    [ObservableProperty]
    private NavItem? _selectedPage;

    /// <summary>True while the app-wide startup connect (see <see cref="ConnectAllAsync"/>) is running.</summary>
    [ObservableProperty] private bool _isConnecting;

    [ObservableProperty] private string _connectingStatusText = "Connecting to Microsoft 365...";

    /// <summary>Device-code sign-in instructions, shown prominently while a connector needs one.</summary>
    [ObservableProperty] private string _deviceCodeMessage = "";

    /// <summary>Just the sign-in code, for the copy button.</summary>
    [ObservableProperty] private string _deviceCode = "";

    private bool _browserOpened;

    public ObservableCollection<NavItem> Pages { get; }

    public MainViewModel(
        DashboardViewModel dashboard,
        GroupsViewModel groups,
        SettingsViewModel settings,
        LogsViewModel logs,
        IStartupConnectionService startup)
    {
        _dashboard = dashboard;
        _groups = groups;
        _startup = startup;

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

    /// <summary>
    /// Connects every registered M365 connector (Graph, Exchange Online, and any future one -
    /// Teams, SharePoint, ... - see IM365Connector) right at app startup, so nobody has to open
    /// Settings to sign in and then Groups to connect separately. Called once from
    /// App.xaml.cs after the main window is shown.
    /// </summary>
    public async Task ConnectAllAsync()
    {
        IsConnecting = true;
        _browserOpened = false;
        DeviceCodeMessage = "";
        DeviceCode = "";
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        try
        {
            await _startup.ConnectAllAsync(
                progress =>
                {
                    void Apply()
                    {
                        ConnectingStatusText = progress.State switch
                        {
                            ConnectorState.Connecting => $"Connecting to {progress.DisplayName}...",
                            ConnectorState.Connected => $"{progress.DisplayName} connected.",
                            ConnectorState.Failed => $"{progress.DisplayName} failed: {progress.Error}",
                            _ => ConnectingStatusText,
                        };
                    }
                    if (dispatcher is not null) dispatcher.Invoke(Apply); else Apply();
                },
                prompt =>
                {
                    void Show()
                    {
                        DeviceCodeMessage = string.IsNullOrEmpty(DeviceCodeMessage) ? prompt : $"{DeviceCodeMessage}\n{prompt}";
                        var code = DeviceCodePrompt.ExtractCode(prompt);
                        if (code is not null)
                            DeviceCode = code;
                        if (!_browserOpened)
                            _browserOpened = DeviceCodePrompt.TryOpenBrowser(prompt);
                    }
                    if (dispatcher is not null) dispatcher.Invoke(Show); else Show();
                });
        }
        finally
        {
            DeviceCodeMessage = "";
            DeviceCode = "";
            IsConnecting = false;
            _dashboard.RefreshCommand.Execute(null);
            _groups.RefreshConnectionState();
        }
    }

    [RelayCommand]
    private void CopyDeviceCode()
    {
        if (string.IsNullOrEmpty(DeviceCode))
            return;
        try
        {
            System.Windows.Clipboard.SetText(DeviceCode);
        }
        catch
        {
            // Clipboard can occasionally be locked by another app; ignore.
        }
    }

    [RelayCommand]
    private void DismissConnecting() => IsConnecting = false;
}
