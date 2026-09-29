using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Controls;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Teams;

namespace M365Manager.ViewModels;

public sealed partial class TeamsViewModel : ObservableObject, ITabbedPage
{
    /// <summary>Which tab the view shows; set by the shell when a Home tile navigates here.</summary>
    [ObservableProperty] private int _selectedTabIndex;

    private readonly ITeamsService _teams;
    private readonly ITeamNamingService _naming;

    [ObservableProperty] private string _taskNumber = "";
    [ObservableProperty] private string _location = "";
    [ObservableProperty] private string _name = "";
    [ObservableProperty] private string _description = "";
    [ObservableProperty] private string _ownerIdentity = "";
    [ObservableProperty] private bool _isPublic;
    [ObservableProperty] private bool _isInternal = true;

    [ObservableProperty] private string _previewName = "";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = "Fill in the fields and click Create team.";

    /// <summary>Guards against an older preview-name lookup overwriting a newer one.</summary>
    private int _previewRequestVersion;

    // --- Teams overview grid (read-only, straight from the SQL cache) ---
    [ObservableProperty] private string _overviewFilterText = "";
    [ObservableProperty] private bool _isOverviewBusy;
    [ObservableProperty] private string _overviewStatusMessage = "";

    // Status severity, so a failure stops looking like a success. Set explicitly at each outcome.
    [ObservableProperty] private AlertSeverity _statusSeverity = AlertSeverity.Info;
    [ObservableProperty] private AlertSeverity _overviewStatusSeverity = AlertSeverity.Info;

    private void Status(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        StatusSeverity = severity;
        StatusMessage = text;
    }

    private void OverviewStatus(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        OverviewStatusSeverity = severity;
        OverviewStatusMessage = text;
    }

    /// <summary>
    /// Row picked in the overview grid. Read-only: Teams has no management actions, so this only
    /// drives the detail card, which carries the columns the grid no longer has room for.
    /// </summary>
    [ObservableProperty] private TeamOverviewRow? _selectedTeam;

    public ObservableCollection<TeamOverviewRow> Teams { get; } = new();

    /// <summary>Filtered view over <see cref="Teams"/> that the grid binds to (and export reads from).</summary>
    public ICollectionView TeamsOverviewView { get; }

    /// <summary>Backs the PowerShell console on this page - see PowerShellConsole.</summary>
    public IPowerShellTranscript Transcript { get; }

    public TeamsViewModel(ITeamsService teams, ITeamNamingService naming, IPowerShellTranscript transcript)
    {
        _teams = teams;
        _naming = naming;
        Transcript = transcript;

        TeamsOverviewView = CollectionViewSource.GetDefaultView(Teams);
        TeamsOverviewView.Filter = FilterTeam;

        _ = RefreshTeamsOverviewAsync();
    }

    partial void OnOverviewFilterTextChanged(string value) => TeamsOverviewView.Refresh();

    private bool FilterTeam(object obj)
    {
        if (string.IsNullOrWhiteSpace(OverviewFilterText))
            return true;
        if (obj is not TeamOverviewRow t)
            return true;

        var term = OverviewFilterText.Trim();
        return Contains(t.DisplayName, term)
            || Contains(t.Alias, term)
            || Contains(t.PrimarySmtpAddress, term)
            || Contains(t.Owners, term)
            || Contains(t.Visibility, term);
    }

    private static bool Contains(string? value, string term)
        => !string.IsNullOrEmpty(value) && value.Contains(term, StringComparison.OrdinalIgnoreCase);

    [RelayCommand]
    private async Task RefreshTeamsOverviewAsync()
    {
        IsOverviewBusy = true;
        OverviewStatusMessage = "Loading teams from SQL...";
        try
        {
            var rows = await _teams.GetTeamsOverviewAsync();
            Teams.Clear();
            foreach (var row in rows)
                Teams.Add(row);

            OverviewStatus($"{Teams.Count} team(s) loaded.", Teams.Count > 0 ? AlertSeverity.Success : AlertSeverity.Info);
        }
        catch (Exception ex)
        {
            OverviewStatus($"Could not load teams: {ex.Message}", AlertSeverity.Error);
        }
        finally
        {
            IsOverviewBusy = false;
        }
    }

    [RelayCommand]
    private void ExportTeamsToExcel()
    {
        var filtered = TeamsOverviewView.Cast<TeamOverviewRow>().ToList();
        if (filtered.Count == 0)
        {
            OverviewStatus("Nothing to export - the (filtered) list is empty.", AlertSeverity.Error);
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = "Teams.xlsx",
        };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            TeamsExportService.ExportToExcel(dialog.FileName, filtered);
            OverviewStatus($"Exported {filtered.Count} team(s) to {dialog.FileName}.", AlertSeverity.Success);
        }
        catch (Exception ex)
        {
            OverviewStatus($"Export failed: {ex.Message}", AlertSeverity.Error);
        }
    }

    partial void OnLocationChanged(string value) => _ = RefreshPreviewAsync();
    partial void OnNameChanged(string value) => _ = RefreshPreviewAsync();

    private async Task RefreshPreviewAsync()
    {
        var version = ++_previewRequestVersion;
        var location = Location;
        var name = Name;

        if (string.IsNullOrWhiteSpace(location) || string.IsNullOrWhiteSpace(name))
        {
            PreviewName = "";
            return;
        }

        try
        {
            var preview = await _naming.BuildDisplayNameAsync(location, name);
            if (version == _previewRequestVersion)
                PreviewName = preview;
        }
        catch
        {
            // Preview is best-effort (e.g. SQL not reachable yet) - leave the last known value.
        }
    }

    [RelayCommand]
    private async Task CreateTeamAsync()
    {
        if (!ValidateInput(out var error))
        {
            Status(error, AlertSeverity.Error);
            return;
        }

        var request = new TeamCreationRequest
        {
            TaskNumber = TaskNumber.Trim(),
            Location = Location.Trim(),
            Name = Name.Trim(),
            Description = Description.Trim(),
            OwnerIdentity = OwnerIdentity.Trim(),
            IsPublic = IsPublic,
            IsInternal = IsInternal,
        };

        IsBusy = true;
        StatusMessage = "Starting team creation...";
        Transcript.BeginOperation($"Create team: {PreviewName}");
        try
        {
            var result = await _teams.CreateTeamAsync(request, msg => StatusMessage = msg);
            StatusSeverity = result.Succeeded ? AlertSeverity.Success : AlertSeverity.Error;
            StatusMessage = result.Succeeded
                ? $"Team '{result.DisplayName}' created successfully."
                : $"Team creation failed: {result.ErrorMessage}";

            if (result.Succeeded)
            {
                ClearForm();
                _ = RefreshTeamsOverviewAsync();
            }
        }
        catch (Exception ex)
        {
            Status($"Team creation failed: {ex.Message}", AlertSeverity.Error);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private bool ValidateInput(out string error)
    {
        if (string.IsNullOrWhiteSpace(TaskNumber)) { error = "Enter the ticket/task number authorizing this change."; return false; }
        if (string.IsNullOrWhiteSpace(Location)) { error = "Enter a location code."; return false; }
        if (string.IsNullOrWhiteSpace(Name)) { error = "Enter a team name."; return false; }
        if (string.IsNullOrWhiteSpace(Description)) { error = "Enter a description."; return false; }
        if (string.IsNullOrWhiteSpace(OwnerIdentity)) { error = "Enter the owner (SamAccountName, UPN or e-mail)."; return false; }
        error = "";
        return true;
    }

    private void ClearForm()
    {
        TaskNumber = "";
        Location = "";
        Name = "";
        Description = "";
        OwnerIdentity = "";
        IsPublic = false;
        IsInternal = true;
        PreviewName = "";
    }
}
