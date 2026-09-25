using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Core.M365;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Services;

namespace M365Manager.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private readonly IStartupConnectionService _startup;
    private readonly IThemeService _theme;
    private readonly ISettingsService _settings;
    private readonly DashboardViewModel _dashboard;
    private readonly GroupsViewModel _groups;
    private readonly TeamsViewModel _teams;
    private readonly SharedMailboxesViewModel _sharedMailboxes;
    private readonly TraceViewModel _trace;

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

    /// <summary>
    /// The in-GUI PowerShell console, hosted once by the shell instead of once per page. It is a
    /// singleton fed centrally from PowerShellHost, so all seven pages were already showing the
    /// same content - seven copies just added 300px to the bottom of each of them.
    /// </summary>
    public IPowerShellTranscript Transcript { get; }

    /// <summary>Whether the console drawer is open. Session-only; it starts closed.</summary>
    [ObservableProperty] private bool _isConsoleOpen;

    public ObservableCollection<NavItem> Pages { get; }

    /// <summary>Grouped view over <see cref="Pages"/>; the sidebar binds to this, not to Pages.</summary>
    public ICollectionView PagesView { get; }

    public MainViewModel(
        DashboardViewModel dashboard,
        GroupsViewModel groups,
        TeamsViewModel teams,
        SharedMailboxesViewModel sharedMailboxes,
        RoomResourcesViewModel roomResources,
        TeamsPoliciesViewModel teamsPolicies,
        NewHireViewModel newHire,
        TraceViewModel trace,
        SettingsViewModel settings,
        LogsViewModel logs,
        IStartupConnectionService startup,
        IThemeService theme,
        ISettingsService settingsService,
        INavigationService navigation,
        IPowerShellTranscript transcript)
    {
        _dashboard = dashboard;
        _groups = groups;
        _teams = teams;
        _sharedMailboxes = sharedMailboxes;
        _trace = trace;
        _startup = startup;
        _theme = theme;
        _settings = settingsService;
        Transcript = transcript;

        // Ordered by how often the work actually happens, not alphabetically: the mailbox and group
        // pages that carry the daily ticket load at the top, then everything Teams - the teams
        // themselves, meeting policies and telephony (New Hire) - together, and the rarely used
        // Trace plus the supporting pages at the bottom. Section drives the sidebar grouping.
        // Glyph values are Segeo MDL2 Assets codes - swap one here if it reads wrong.
        Pages = new ObservableCollection<NavItem>
        {
            new() { Title = "Home",             Section = "",                 Glyph = "", ViewModel = dashboard },

            new() { Title = "Groups",           Section = "DAILY OPERATIONS", Glyph = "", ViewModel = groups },
            new() { Title = "Shared Mailboxes", Section = "DAILY OPERATIONS", Glyph = "", ViewModel = sharedMailboxes },
            new() { Title = "Rooms & Resources",Section = "DAILY OPERATIONS", Glyph = "", ViewModel = roomResources },

            new() { Title = "Teams",            Section = "TEAMS",            Glyph = "", ViewModel = teams },
            new() { Title = "Teams Policies",   Section = "TEAMS",            Glyph = "", ViewModel = teamsPolicies },
            new() { Title = "New Hire",         Section = "TEAMS",            Glyph = "", ViewModel = newHire },

            new() { Title = "Trace Guests",     Section = "MORE",             Glyph = "", ViewModel = trace },

            new() { Title = "Logs",             Section = "SYSTEM",           Glyph = "", ViewModel = logs },
            new() { Title = "Settings",         Section = "SYSTEM",           Glyph = "", ViewModel = settings },
        };

        // Grouped view so the rail can draw section captions. Grouping by a plain property keeps
        // the collection itself flat, so selection still moves across the whole rail with the
        // arrow keys rather than being trapped inside one group.
        PagesView = CollectionViewSource.GetDefaultView(Pages);
        PagesView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(NavItem.Section)));

        SelectedPage = Pages[0];
        RefreshThemeLabel();
        CheckSettingsLoad();

        navigation.Requested += OnNavigationRequested;
    }

    [RelayCommand]
    private void ToggleConsole() => IsConsoleOpen = !IsConsoleOpen;

    /// <summary>
    /// Ctrl+1..0 jumps straight to the nth page in the rail. Ticket work means hopping between
    /// Groups, Shared Mailboxes and Teams all day; reaching for the mouse each time adds up.
    /// Out-of-range indexes are ignored so the binding cannot throw.
    /// </summary>
    [RelayCommand]
    private void GoToPageByIndex(string? oneBasedIndex)
    {
        if (!int.TryParse(oneBasedIndex, out var position))
            return;

        // Ctrl+0 is the tenth entry, matching how browsers number their tab shortcuts.
        var index = position == 0 ? 9 : position - 1;
        if (index >= 0 && index < Pages.Count)
            SelectedPage = Pages[index];
    }

    /// <summary>
    /// Handles a page asking to move - currently only the Home tiles. Unknown titles are ignored
    /// rather than throwing: a tile with a typo should do nothing, not take the app down.
    /// </summary>
    private void OnNavigationRequested(NavigationRequest request)
    {
        var target = Pages.FirstOrDefault(p =>
            string.Equals(p.Title, request.PageTitle, StringComparison.Ordinal));
        if (target is null)
            return;

        SelectedPage = target;

        // Set the tab after the page, so the view is bound to this view model by the time the
        // index arrives.
        if (request.TabIndex is { } index && target.ViewModel is ITabbedPage tabbed)
            tabbed.SelectedTabIndex = index;
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

                        // Each connector runs its own device-code flow and so needs its own browser
                        // visit and its own code. Without this reset the banner accumulated every
                        // connector's prompt - showing a stale code above the current one - and the
                        // browser opened only for whichever connector happened to prompt first,
                        // leaving the rest to be completed by hand.
                        if (progress.State == ConnectorState.Connecting)
                        {
                            DeviceCodeMessage = "";
                            DeviceCode = "";
                            _browserOpened = false;
                        }
                    }
                    if (dispatcher is not null) dispatcher.Invoke(Apply); else Apply();
                },
                prompt =>
                {
                    void Show()
                    {
                        // A single connector can emit several lines (a headline, then the code), so
                        // these still accumulate - they are just cleared between connectors above.
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
            _sharedMailboxes.RefreshAvailableDomains();
            _trace.RefreshAfterConnect();
        }
    }

    // ----- Settings load warning -----

    /// <summary>
    /// Non-empty when the settings file could not be read in full. Shown as a persistent strip
    /// under the header rather than a toast: saving anything while this is up would write defaults
    /// over the parts that failed to load, so it has to stay visible until the app is restarted
    /// with a good file.
    /// </summary>
    [ObservableProperty] private string _settingsLoadWarning = "";

    private void CheckSettingsLoad()
    {
        if (string.IsNullOrEmpty(_settings.LastLoadError))
            return;

        SettingsLoadWarning = _settings.QuarantinedFilePath is { Length: > 0 } backup
            ? $"{_settings.LastLoadError} A copy of the original file was kept at {backup}."
            : _settings.LastLoadError;
    }

    [RelayCommand]
    private void DismissSettingsWarning() => SettingsLoadWarning = "";

    // ----- Theme -----

    /// <summary>Label on the sidebar's theme button, e.g. "Theme: System".</summary>
    [ObservableProperty] private string _themeLabel = "";

    /// <summary>Segoe MDL2 glyph matching the theme currently on screen.</summary>
    [ObservableProperty] private string _themeGlyph = "";

    /// <summary>
    /// Cycles System -> Light -> Dark -> System. Three states rather than a plain on/off switch,
    /// because "follow Windows" is the sensible default and a two-way toggle cannot express it.
    /// </summary>
    [RelayCommand]
    private void CycleTheme()
    {
        var next = _theme.Preference switch
        {
            AppTheme.System => AppTheme.Light,
            AppTheme.Light => AppTheme.Dark,
            _ => AppTheme.System,
        };

        _theme.SetPreference(next);
        RefreshThemeLabel();
    }

    private void RefreshThemeLabel()
    {
        ThemeLabel = _theme.Preference switch
        {
            AppTheme.Light => "Theme: Light",
            AppTheme.Dark => "Theme: Dark",
            _ => "Theme: System",
        };

        // Reflects what is actually on screen, which for "System" depends on Windows.
        ThemeGlyph = _theme.Effective == AppTheme.Dark ? "" : "";
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
