using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Core.Trace;

namespace M365Manager.ViewModels;

public sealed partial class TraceViewModel : ObservableObject
{
    private readonly ITraceGuestService _guests;
    private readonly ISettingsService _settings;

    // --- Invite form ---
    [ObservableProperty] private string _displayName = "";
    [ObservableProperty] private string _email = "";
    [ObservableProperty] private string _taskNumber = "";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = "Enter the guest's name and e-mail address, then click Invite.";

    /// <summary>Redemption link from the last invitation - handy when the mail doesn't arrive.</summary>
    [ObservableProperty] private string _lastRedeemUrl = "";

    // --- Overview ---
    [ObservableProperty] private string _filterText = "";
    [ObservableProperty] private bool _isOverviewBusy;
    [ObservableProperty] private string _overviewStatusMessage = "";

    public ObservableCollection<TraceGuestRow> Guests { get; } = new();
    public ICollectionView GuestsView { get; }

    /// <summary>Shown on the page so it's obvious which company name new guests are stamped with.</summary>
    public string CompanyName => _settings.Current.Trace.CompanyName;
    public string RedirectUrl => _settings.Current.Trace.RedirectUrl;

    /// <summary>
    /// Backs the console on this page. Trace is pure Graph - without the console showing Graph
    /// requests, inviting a guest looked as though nothing happened at all.
    /// </summary>
    public IPowerShellTranscript Transcript { get; }

    public TraceViewModel(ITraceGuestService guests, ISettingsService settings, IPowerShellTranscript transcript)
    {
        _guests = guests;
        _settings = settings;
        Transcript = transcript;

        GuestsView = CollectionViewSource.GetDefaultView(Guests);
        GuestsView.Filter = FilterRow;
    }

    /// <summary>
    /// Called by MainViewModel once the app-wide sign-in has finished - the guest list needs Graph,
    /// which isn't connected yet while the constructor runs.
    /// </summary>
    public void RefreshAfterConnect()
    {
        OnPropertyChanged(nameof(CompanyName));
        OnPropertyChanged(nameof(RedirectUrl));
        _ = RefreshGuestsAsync();
    }

    partial void OnFilterTextChanged(string value) => GuestsView.Refresh();

    private bool FilterRow(object obj)
    {
        if (string.IsNullOrWhiteSpace(FilterText))
            return true;
        if (obj is not TraceGuestRow g)
            return true;

        var term = FilterText.Trim();
        return Contains(g.DisplayName, term)
            || Contains(g.Mail, term)
            || Contains(g.UserPrincipalName, term)
            || Contains(g.StatusLabel, term);
    }

    private static bool Contains(string? value, string term)
        => !string.IsNullOrEmpty(value) && value.Contains(term, StringComparison.OrdinalIgnoreCase);

    [RelayCommand]
    private async Task InviteAsync()
    {
        if (string.IsNullOrWhiteSpace(DisplayName))
        {
            StatusMessage = "Enter the guest's display name.";
            return;
        }
        if (string.IsNullOrWhiteSpace(Email) || !Email.Contains('@'))
        {
            StatusMessage = "Enter a valid e-mail address.";
            return;
        }

        var request = new TraceGuestInvitationRequest
        {
            DisplayName = DisplayName.Trim(),
            Email = Email.Trim(),
            TaskNumber = TaskNumber.Trim(),
        };

        IsBusy = true;
        LastRedeemUrl = "";
        StatusMessage = "Sending invitation...";
        Transcript.BeginOperation($"Invite Trace guest: {Email.Trim()}");
        try
        {
            var result = await _guests.InviteAsync(request, msg => StatusMessage = msg);

            if (result.Succeeded)
            {
                StatusMessage = result.Warning is { Length: > 0 }
                    ? $"{result.DisplayName} invited - but note: {result.Warning}"
                    : $"{result.DisplayName} <{result.Email}> invited and stamped with company name '{CompanyName}'.";

                LastRedeemUrl = result.RedeemUrl ?? "";
                DisplayName = "";
                Email = "";
                TaskNumber = "";
                _ = RefreshGuestsAsync();
            }
            else
            {
                StatusMessage = $"Invitation failed: {result.ErrorMessage}";
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Invitation failed: {ErrorText.Describe(ex)}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task RefreshGuestsAsync()
    {
        IsOverviewBusy = true;
        OverviewStatusMessage = "Loading guest accounts from Entra...";
        try
        {
            var rows = await _guests.GetGuestsAsync();
            Guests.Clear();
            foreach (var row in rows)
                Guests.Add(row);

            OverviewStatusMessage = Guests.Count == 0
                ? $"No guest accounts found with company name '{CompanyName}'."
                : $"{Guests.Count} guest account(s) with company name '{CompanyName}'.";
        }
        catch (Exception ex)
        {
            OverviewStatusMessage = $"Could not load guest accounts: {ErrorText.Describe(ex)}";
        }
        finally
        {
            IsOverviewBusy = false;
        }
    }

    [RelayCommand]
    private void CopyRedeemUrl()
    {
        if (string.IsNullOrEmpty(LastRedeemUrl))
            return;
        try
        {
            System.Windows.Clipboard.SetText(LastRedeemUrl);
            StatusMessage = "Redemption link copied to the clipboard.";
        }
        catch
        {
            // Clipboard can be locked by another app; ignore.
        }
    }

    [RelayCommand]
    private void ExportToExcel()
    {
        var filtered = GuestsView.Cast<TraceGuestRow>().ToList();
        if (filtered.Count == 0)
        {
            OverviewStatusMessage = "Nothing to export - the (filtered) list is empty.";
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = "TraceGuests.xlsx",
        };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            TraceGuestExportService.ExportToExcel(dialog.FileName, filtered);
            OverviewStatusMessage = $"Exported {filtered.Count} guest(s) to {dialog.FileName}.";
        }
        catch (Exception ex)
        {
            OverviewStatusMessage = $"Export failed: {ErrorText.Describe(ex)}";
        }
    }
}
