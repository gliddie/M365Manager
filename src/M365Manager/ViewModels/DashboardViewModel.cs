using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Controls;
using M365Manager.Core.M365;
using M365Manager.Core.Settings;
using M365Manager.Data.Logging;
using M365Manager.Services;

namespace M365Manager.ViewModels;

/// <summary>
/// The home page. Was two status cards and a Refresh button - the only screen in the app without a
/// task, while also being the one everybody lands on.
///
/// Now it is the ticket-driven entry point: a compact connection line, shortcuts straight into the
/// forms the daily tickets ask for, and the last few log entries so a run that went wrong is
/// visible without opening the Logs page.
/// </summary>
public sealed partial class DashboardViewModel : ObservableObject
{
    private readonly ISettingsService _settings;
    private readonly IM365AuthService _auth;
    private readonly ILogService _log;
    private readonly INavigationService _navigation;

    [ObservableProperty] private string _sqlStatus = "";
    [ObservableProperty] private string _m365Status = "";
    [ObservableProperty] private bool _sqlConfigured;
    [ObservableProperty] private bool _signedIn;

    [ObservableProperty] private string _recentStatus = "";
    [ObservableProperty] private AlertSeverity _recentStatusSeverity = AlertSeverity.Info;

    /// <summary>The last handful of log entries, newest first.</summary>
    public ObservableCollection<LogEntry> RecentEntries { get; } = new();

    public DashboardViewModel(ISettingsService settings, IM365AuthService auth, ILogService log, INavigationService navigation)
    {
        _settings = settings;
        _auth = auth;
        _log = log;
        _navigation = navigation;
        Refresh();
    }

    [RelayCommand]
    private void Refresh()
    {
        var sql = _settings.Current.Sql;
        SqlConfigured = !string.IsNullOrWhiteSpace(sql.Server);
        SqlStatus = SqlConfigured
            ? $"{sql.Server}  •  {sql.Database}"
            : "Not configured - open Settings.";

        SignedIn = _auth.IsSignedIn;
        M365Status = SignedIn
            ? $"Signed in as {_auth.CurrentUser!.Upn}"
            : "Not signed in.";

        _ = LoadRecentAsync();
    }

    private async Task LoadRecentAsync()
    {
        if (!SqlConfigured)
        {
            RecentEntries.Clear();
            RecentStatus = "Recent activity needs the SQL connection - configure it under Settings.";
            RecentStatusSeverity = AlertSeverity.Info;
            return;
        }

        try
        {
            var entries = await _log.QueryAsync(new LogQuery { MaxResults = 8 });
            RecentEntries.Clear();
            foreach (var entry in entries)
                RecentEntries.Add(entry);

            RecentStatus = RecentEntries.Count == 0 ? "Nothing logged yet." : "";
            RecentStatusSeverity = AlertSeverity.Info;
        }
        catch (Exception ex)
        {
            RecentEntries.Clear();
            RecentStatus = $"Could not read the log: {ErrorText.Describe(ex)}";
            RecentStatusSeverity = AlertSeverity.Error;
        }
    }

    /// <summary>
    /// Jumps to a page, and to a specific tab on it. The parameter is "Page" or "Page|tabIndex" so
    /// the tiles stay declarative in XAML rather than each needing its own command.
    /// </summary>
    [RelayCommand]
    private void Open(string? target)
    {
        if (string.IsNullOrWhiteSpace(target))
            return;

        var parts = target.Split('|', 2);
        var tab = parts.Length == 2 && int.TryParse(parts[1], out var index) ? index : (int?)null;
        _navigation.GoTo(parts[0], tab);
    }

    [RelayCommand]
    private void OpenLogs() => _navigation.GoTo("Logs");
}
