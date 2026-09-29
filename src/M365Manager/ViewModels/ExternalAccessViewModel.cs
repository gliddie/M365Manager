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

/// <summary>
/// Teams external access (federation): the tenant-wide allow list of partner domains. Its own page
/// in the TEAMS section - it started as a tab on the Teams page, but that page creates and looks up
/// individual teams, and a tenant-wide setting sitting among them was misleading.
///
/// Live from Teams rather than from the SQL cache: it is one list, and the add/remove decisions
/// depend on what it says right now.
/// </summary>
public sealed partial class ExternalAccessViewModel : ObservableObject, IActivatablePage
{
    private readonly ITeamsFederationService _federation;
    private readonly IDialogService _dialogs;

    /// <summary>Backs the PowerShell console - see PowerShellConsole.</summary>
    public IPowerShellTranscript Transcript { get; }

    public ExternalAccessViewModel(ITeamsFederationService federation, IDialogService dialogs, IPowerShellTranscript transcript)
    {
        _federation = federation;
        _dialogs = dialogs;
        Transcript = transcript;

        FederationDomainsView = new ListCollectionView(FederationDomains) { Filter = FilterDomain };
    }

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

    /// <summary>Loads the list the first time the page is opened - it needs the Teams connection, so not at startup.</summary>
    public void OnActivated()
    {
        if (!HasFederationState && !IsFederationBusy)
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
