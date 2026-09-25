using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Controls;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Teams;
using M365Manager.Services;

namespace M365Manager.ViewModels;

public sealed partial class TeamsViewModel : ObservableObject, ITabbedPage
{
    /// <summary>Which tab the view shows; set by the shell when a Home tile navigates here.</summary>
    [ObservableProperty] private int _selectedTabIndex;

    private readonly ITeamsService _teams;
    private readonly ITeamNamingService _naming;
    private readonly ITeamsFederationService _federation;
    private readonly IDialogService _dialogs;

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

    public TeamsViewModel(ITeamsService teams, ITeamNamingService naming, ITeamsFederationService federation,
        IDialogService dialogs, IPowerShellTranscript transcript)
    {
        _teams = teams;
        _naming = naming;
        _federation = federation;
        _dialogs = dialogs;
        Transcript = transcript;

        TeamsOverviewView = CollectionViewSource.GetDefaultView(Teams);
        TeamsOverviewView.Filter = FilterTeam;

        // Its own view object: GetDefaultView would hand the same view to anything else bound to
        // this collection, and the filter belongs to this tab only.
        FederationDomainsView = new ListCollectionView(FederationDomains) { Filter = FilterDomain };

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

    // ----- External access (federation) -----
    //
    // Live from Teams rather than from the SQL cache like the overview: it is one tenant-wide
    // list, and the add/remove decisions depend on what it says right now.

    /// <summary>Index of the "External access" tab in the view's TabStrip.</summary>
    private const int FederationTabIndex = 2;

    public ObservableCollection<string> FederationDomains { get; } = new();
    public ICollectionView FederationDomainsView { get; }

    [ObservableProperty] private string _federationFilterText = "";
    [ObservableProperty] private string? _selectedFederationDomain;
    [ObservableProperty] private string _federationTaskNumber = "";
    [ObservableProperty] private string _federationDomainsInput = "";
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanSubmitFederationAdd))] private bool _isFederationBusy;
    [ObservableProperty] private bool _hasFederationState;

    /// <summary>Adding is only offered in closed federation - see FederationMode.AllowAllKnownDomains.</summary>
    [ObservableProperty, NotifyPropertyChangedFor(nameof(CanSubmitFederationAdd))] private bool _canAddFederationDomains;

    public bool CanSubmitFederationAdd => CanAddFederationDomains && !IsFederationBusy;

    [ObservableProperty] private string _federationModeText = "";
    [ObservableProperty] private AlertSeverity _federationModeSeverity = AlertSeverity.Info;
    [ObservableProperty] private string _federationListHint = "Loading the allowed domains from Teams...";
    [ObservableProperty] private string _federationStatusMessage = "";
    [ObservableProperty] private AlertSeverity _federationStatusSeverity = AlertSeverity.Info;

    private void FederationStatus(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        FederationStatusSeverity = severity;
        FederationStatusMessage = text;
    }

    /// <summary>Loads the list the first time the tab is opened - it needs the Teams connection, so not at startup.</summary>
    partial void OnSelectedTabIndexChanged(int value)
    {
        if (value == FederationTabIndex && !HasFederationState && !IsFederationBusy)
            _ = RefreshFederationAsync();
    }

    partial void OnFederationFilterTextChanged(string value) => FederationDomainsView.Refresh();

    private bool FilterDomain(object obj)
        => string.IsNullOrWhiteSpace(FederationFilterText)
           || (obj is string domain && domain.Contains(FederationFilterText.Trim(), StringComparison.OrdinalIgnoreCase));

    [RelayCommand]
    private async Task RefreshFederationAsync()
    {
        IsFederationBusy = true;
        FederationListHint = "Loading the allowed domains from Teams...";
        Transcript.BeginOperation("Read Teams federation configuration");
        try
        {
            ApplyFederationState(await _federation.GetStateAsync());
        }
        catch (Exception ex)
        {
            HasFederationState = false;
            CanAddFederationDomains = false;
            FederationModeText = "";
            FederationListHint = "Could not read the federation configuration - check the Teams connection on the Dashboard.";
            FederationStatus($"Could not load the federation configuration: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsFederationBusy = false;
        }
    }

    private void ApplyFederationState(FederationState state)
    {
        var selected = SelectedFederationDomain;
        FederationDomains.Clear();
        foreach (var domain in state.AllowedDomains)
            FederationDomains.Add(domain);
        SelectedFederationDomain = selected is not null && FederationDomains.Contains(selected) ? selected : null;
        HasFederationState = true;

        if (state.Mode == FederationMode.AllowAllKnownDomains)
        {
            CanAddFederationDomains = false;
            FederationModeSeverity = AlertSeverity.Warning;
            FederationModeText = "Open federation: Teams users can reach every external domain except blocked ones. "
                                 + "Adding a domain here would switch the tenant to an allow list and cut off everyone else, "
                                 + "so adding is disabled. Change the mode in the Teams admin center if that is intended.";
            FederationListHint = "There is no allow list in open federation.";
            return;
        }

        CanAddFederationDomains = true;
        if (!state.AllowFederatedUsers)
        {
            FederationModeSeverity = AlertSeverity.Warning;
            FederationModeText = "External access is switched off for the whole tenant (AllowFederatedUsers = False). "
                                 + "The domains below have no effect until it is switched on in the Teams admin center.";
        }
        else
        {
            FederationModeSeverity = AlertSeverity.Info;
            FederationModeText = $"Closed federation: Teams users can chat, call and meet only with the {state.AllowedDomains.Count} "
                                 + "partner domain(s) below. Guest access (B2B) is separate and not affected.";
        }
        FederationListHint = "The allow list is empty - no external domain is reachable through federation.";
    }

    [RelayCommand]
    private async Task AddFederationDomainsAsync()
    {
        if (string.IsNullOrWhiteSpace(FederationTaskNumber))
        {
            FederationStatus("Enter the ticket/task number authorizing this change.", AlertSeverity.Error);
            return;
        }

        var (domains, invalid) = TeamsFederationService.ParseDomains(FederationDomainsInput);
        if (invalid.Count > 0)
        {
            FederationStatus($"Not a domain: {string.Join(", ", invalid)}. Enter domains like partner.com, one per line or comma-separated.", AlertSeverity.Error);
            return;
        }
        if (domains.Count == 0)
        {
            FederationStatus("Enter at least one partner domain.", AlertSeverity.Error);
            return;
        }

        IsFederationBusy = true;
        FederationStatus($"Adding {string.Join(", ", domains)}...");
        Transcript.BeginOperation($"Add federated domain(s): {string.Join(", ", domains)}");
        try
        {
            var result = await _federation.AddDomainsAsync(FederationTaskNumber.Trim(), domains);
            if (!result.Succeeded)
            {
                FederationStatus($"Adding failed: {result.ErrorMessage}", AlertSeverity.Error);
                return;
            }

            var skipped = result.Skipped.Count == 0 ? "" : $" Already allowed: {string.Join(", ", result.Skipped)}.";
            FederationStatus(result.Changed.Count == 0
                    ? $"Nothing to add.{skipped}"
                    : $"Added {string.Join(", ", result.Changed)}.{skipped} It can take a few hours until Teams applies the change everywhere.",
                result.Changed.Count == 0 ? AlertSeverity.Info : AlertSeverity.Success);

            FederationDomainsInput = "";
            FederationTaskNumber = "";
            ApplyFederationState(await _federation.GetStateAsync());
        }
        catch (Exception ex)
        {
            FederationStatus($"Adding failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsFederationBusy = false;
        }
    }

    [RelayCommand]
    private async Task RemoveFederationDomainAsync()
    {
        if (SelectedFederationDomain is not { } domain)
        {
            FederationStatus("Select the domain to remove in the list.", AlertSeverity.Error);
            return;
        }
        if (string.IsNullOrWhiteSpace(FederationTaskNumber))
        {
            FederationStatus("Enter the ticket/task number authorizing this change.", AlertSeverity.Error);
            return;
        }

        var last = FederationDomains.Count == 1
            ? "\n\nIt is the last domain on the list - afterwards no external domain is reachable through federation."
            : "";
        if (!_dialogs.ConfirmDestructive(
                $"Remove {domain}?",
                $"Teams users can no longer chat, call or meet with people from {domain} through external access.",
                "Remove domain",
                "Existing chats stay visible but read-only. Guest accounts (B2B) from this partner are not affected." + last))
            return;

        IsFederationBusy = true;
        FederationStatus($"Removing {domain}...");
        Transcript.BeginOperation($"Remove federated domain: {domain}");
        try
        {
            var result = await _federation.RemoveDomainAsync(FederationTaskNumber.Trim(), domain);
            if (!result.Succeeded)
            {
                FederationStatus($"Removing failed: {result.ErrorMessage}", AlertSeverity.Error);
                return;
            }

            FederationStatus(result.Changed.Count == 0
                    ? $"{domain} was no longer on the list - nothing to remove."
                    : $"Removed {domain}.",
                result.Changed.Count == 0 ? AlertSeverity.Info : AlertSeverity.Success);

            FederationTaskNumber = "";
            ApplyFederationState(await _federation.GetStateAsync());
        }
        catch (Exception ex)
        {
            FederationStatus($"Removing failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsFederationBusy = false;
        }
    }
}
