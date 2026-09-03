using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Core.SharedMailboxes;

namespace M365Manager.ViewModels;

public sealed partial class SharedMailboxesViewModel : ObservableObject
{
    private readonly ISharedMailboxService _sharedMailboxes;
    private readonly ISharedMailboxNamingService _naming;

    // --- Create form ---
    [ObservableProperty] private string _taskNumber = "";
    [ObservableProperty] private string _location = "";
    [ObservableProperty] private string _name = "";
    [ObservableProperty] private string? _domain;
    [ObservableProperty] private string _ownerIdentities = "";
    [ObservableProperty] private string _editorMembers = "";
    [ObservableProperty] private string _authorMembers = "";
    [ObservableProperty] private string _readerMembers = "";
    [ObservableProperty] private bool _allowExternalSenders;

    [ObservableProperty] private string _previewName = "";
    [ObservableProperty] private string _previewAddress = "";

    /// <summary>True once the operator has typed in the corresponding preview field by hand.</summary>
    [ObservableProperty] private bool _isPreviewNameEdited;
    [ObservableProperty] private bool _isPreviewAddressEdited;

    /// <summary>Guards the edited-flags while the preview is being filled in programmatically.</summary>
    private bool _applyingPreview;

    private void ApplyPreview(string name, string address)
    {
        _applyingPreview = true;
        try
        {
            PreviewName = name;
            PreviewAddress = address;
        }
        finally
        {
            _applyingPreview = false;
        }
    }
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = "Fill in the fields and click Create Shared Mailbox.";

    public ObservableCollection<string> AvailableDomains { get; } = new();

    /// <summary>Guards against an older preview lookup overwriting a newer one.</summary>
    private int _previewRequestVersion;

    // --- Change owner ---
    [ObservableProperty] private string _changeOwnerTaskNumber = "";
    [ObservableProperty] private string _changeOwnerMailboxAddress = "";
    [ObservableProperty] private string _changeOwnerIdentities = "";
    [ObservableProperty] private OwnerChangeMode _changeOwnerMode = OwnerChangeMode.Add;
    [ObservableProperty] private bool _isChangingOwner;
    [ObservableProperty] private string _changeOwnerStatusMessage = "";

    // --- Overview grid (read-only, straight from the SQL cache) ---
    [ObservableProperty] private string _overviewFilterText = "";
    [ObservableProperty] private bool _isOverviewBusy;
    [ObservableProperty] private string _overviewStatusMessage = "";

    public ObservableCollection<SharedMailboxOverviewRow> SharedMailboxes { get; } = new();

    /// <summary>Filtered view over <see cref="SharedMailboxes"/> that the grid binds to (and export reads from).</summary>
    public ICollectionView SharedMailboxesOverviewView { get; }

    /// <summary>Backs the PowerShell console on this page - see PowerShellConsole.</summary>
    public IPowerShellTranscript Transcript { get; }

    public SharedMailboxesViewModel(ISharedMailboxService sharedMailboxes, ISharedMailboxNamingService naming, ISettingsService settings, IPowerShellTranscript transcript)
    {
        _sharedMailboxes = sharedMailboxes;
        _naming = naming;
        Transcript = transcript;

        SharedMailboxesOverviewView = CollectionViewSource.GetDefaultView(SharedMailboxes);
        SharedMailboxesOverviewView.Filter = FilterRow;

        // Pre-fill immediately from Settings > Microsoft 365 > Default mail domain, rather than
        // waiting on the async Graph /domains lookup (which needs sign-in to have completed) -
        // covers the common case (one domain used almost always) without the field sitting empty
        // in the meantime. GetAvailableDomainsAsync also puts this same domain first once the
        // real list loads, and LoadAvailableDomainsAsync below only overwrites Domain if it's
        // still blank, so this pre-filled value is kept.
        // Deliberately M365.DefaultMailDomain, NOT M365.Domain - the latter is the UPN/sign-in
        // domain (e.g. "global.ul.com"), which is usually different from the mail/SMTP domain
        // used for mailbox addresses (e.g. "ul.com").
        var defaultDomain = settings.Current.M365.DefaultMailDomain.Trim();
        if (defaultDomain.Length > 0)
            Domain = defaultDomain;

        _ = LoadAvailableDomainsAsync();
        _ = RefreshOverviewAsync();
    }

    /// <summary>
    /// Reloads the address-domain list from Graph. Called by MainViewModel after the app-wide
    /// startup connect finishes - the constructor's own load attempt (below) normally runs before
    /// sign-in completes and silently finds nothing, so without this the dropdown stays empty
    /// until something else happens to call it again.
    /// </summary>
    public void RefreshAvailableDomains() => _ = LoadAvailableDomainsAsync();

    private async Task LoadAvailableDomainsAsync()
    {
        try
        {
            var domains = await _sharedMailboxes.GetAvailableDomainsAsync();
            AvailableDomains.Clear();
            foreach (var d in domains)
                AvailableDomains.Add(d);
            if (string.IsNullOrWhiteSpace(Domain))
                Domain = AvailableDomains.FirstOrDefault();
        }
        catch
        {
            // Domain list is best-effort (e.g. not signed in yet) - the ComboBox stays empty/editable.
        }
    }

    partial void OnOverviewFilterTextChanged(string value) => SharedMailboxesOverviewView.Refresh();

    private bool FilterRow(object obj)
    {
        if (string.IsNullOrWhiteSpace(OverviewFilterText))
            return true;
        if (obj is not SharedMailboxOverviewRow r)
            return true;

        var term = OverviewFilterText.Trim();
        return Contains(r.DisplayName, term)
            || Contains(r.Alias, term)
            || Contains(r.PrimarySmtpAddress, term)
            || Contains(r.Owners, term);
    }

    private static bool Contains(string? value, string term)
        => !string.IsNullOrEmpty(value) && value.Contains(term, StringComparison.OrdinalIgnoreCase);

    [RelayCommand]
    private async Task RefreshOverviewAsync()
    {
        IsOverviewBusy = true;
        OverviewStatusMessage = "Loading shared mailboxes from SQL...";
        try
        {
            var rows = await _sharedMailboxes.GetSharedMailboxesOverviewAsync();
            SharedMailboxes.Clear();
            foreach (var row in rows)
                SharedMailboxes.Add(row);

            OverviewStatusMessage = $"{SharedMailboxes.Count} shared mailbox(es) loaded.";
        }
        catch (Exception ex)
        {
            OverviewStatusMessage = $"Could not load shared mailboxes: {ex.Message}";
        }
        finally
        {
            IsOverviewBusy = false;
        }
    }

    [RelayCommand]
    private void ExportToExcel()
    {
        var filtered = SharedMailboxesOverviewView.Cast<SharedMailboxOverviewRow>().ToList();
        if (filtered.Count == 0)
        {
            OverviewStatusMessage = "Nothing to export - the (filtered) list is empty.";
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = "SharedMailboxes.xlsx",
        };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            SharedMailboxExportService.ExportToExcel(dialog.FileName, filtered);
            OverviewStatusMessage = $"Exported {filtered.Count} shared mailbox(es) to {dialog.FileName}.";
        }
        catch (Exception ex)
        {
            OverviewStatusMessage = $"Export failed: {ex.Message}";
        }
    }

    partial void OnLocationChanged(string value) => _ = RefreshPreviewAsync();
    partial void OnNameChanged(string value) => _ = RefreshPreviewAsync();
    partial void OnDomainChanged(string? value) => _ = RefreshPreviewAsync();

    // PreviewName/PreviewAddress are editable: the convention fills them in, and the operator can
    // correct anything it can't derive (the legacy dialog worked the same way). Once a field has
    // been edited by hand it stops being overwritten, so typing on in the Name field doesn't
    // silently discard the correction.
    partial void OnPreviewNameChanged(string value)
    {
        if (!_applyingPreview)
            IsPreviewNameEdited = true;
    }

    partial void OnPreviewAddressChanged(string value)
    {
        if (!_applyingPreview)
            IsPreviewAddressEdited = true;
    }

    /// <summary>Drops the manual corrections and goes back to what the naming convention produces.</summary>
    [RelayCommand]
    private void ResetPreview()
    {
        IsPreviewNameEdited = false;
        IsPreviewAddressEdited = false;
        _ = RefreshPreviewAsync();
    }

    private async Task RefreshPreviewAsync()
    {
        var version = ++_previewRequestVersion;
        var location = Location;
        var name = Name;
        var domain = Domain;

        if (string.IsNullOrWhiteSpace(location) || string.IsNullOrWhiteSpace(name))
        {
            ApplyPreview(
                IsPreviewNameEdited ? PreviewName : "",
                IsPreviewAddressEdited ? PreviewAddress : "");
            return;
        }

        try
        {
            var (displayName, localPart) = await _naming.BuildNameAsync(location, name);
            if (version != _previewRequestVersion)
                return;

            ApplyPreview(
                IsPreviewNameEdited ? PreviewName : displayName,
                IsPreviewAddressEdited || string.IsNullOrWhiteSpace(domain)
                    ? PreviewAddress
                    : _naming.BuildAddress(localPart, domain));
        }
        catch
        {
            // Preview is best-effort (e.g. SQL not reachable yet) - leave the last known value.
        }
    }

    [RelayCommand]
    private async Task CreateSharedMailboxAsync()
    {
        if (!ValidateInput(out var error))
        {
            StatusMessage = error;
            return;
        }

        var request = new SharedMailboxCreationRequest
        {
            TaskNumber = TaskNumber.Trim(),
            Location = Location.Trim(),
            Name = Name.Trim(),
            Domain = Domain!.Trim(),
            OwnerIdentities = OwnerIdentities.Trim(),
            EditorMembers = EditorMembers.Trim(),
            AuthorMembers = AuthorMembers.Trim(),
            ReaderMembers = ReaderMembers.Trim(),
            AllowExternalSenders = AllowExternalSenders,
            // Whatever is in the preview fields wins - it's either the computed value or the
            // operator's correction of it.
            DisplayNameOverride = PreviewName.Trim(),
            AddressOverride = PreviewAddress.Trim(),
        };

        IsBusy = true;
        StatusMessage = "Starting shared mailbox creation...";
        Transcript.BeginOperation($"Create shared mailbox: {PreviewName.Trim()}");
        try
        {
            var result = await _sharedMailboxes.CreateSharedMailboxAsync(request, msg => StatusMessage = msg);
            StatusMessage = result.Succeeded
                ? $"Shared mailbox '{result.DisplayName}' ({result.PrimarySmtpAddress}) created successfully."
                  + (result.WarningMessage is null ? "" : $" WARNING: {result.WarningMessage}")
                : $"Shared mailbox creation failed: {result.ErrorMessage}";

            if (result.Succeeded)
            {
                ClearForm();
                _ = RefreshOverviewAsync();
            }
        }
        catch (Exception ex)
        {
            StatusMessage = $"Shared mailbox creation failed: {ex.Message}";
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
        if (string.IsNullOrWhiteSpace(Name)) { error = "Enter a mailbox name."; return false; }
        if (string.IsNullOrWhiteSpace(Domain)) { error = "Select an address domain."; return false; }
        if (string.IsNullOrWhiteSpace(OwnerIdentities)) { error = "Enter at least one owner (SamAccountName, UPN or e-mail)."; return false; }
        if (string.IsNullOrWhiteSpace(PreviewName)) { error = "The mailbox name is empty - check the name field below the preview."; return false; }
        if (string.IsNullOrWhiteSpace(PreviewAddress) || !PreviewAddress.Contains('@')) { error = "The mailbox address is not a valid e-mail address."; return false; }
        error = "";
        return true;
    }

    private void ClearForm()
    {
        TaskNumber = "";
        Location = "";
        Name = "";
        OwnerIdentities = "";
        EditorMembers = "";
        AuthorMembers = "";
        ReaderMembers = "";
        AllowExternalSenders = false;
        IsPreviewNameEdited = false;
        IsPreviewAddressEdited = false;
        ApplyPreview("", "");
    }

    [RelayCommand]
    private async Task ChangeOwnerAsync()
    {
        if (string.IsNullOrWhiteSpace(ChangeOwnerTaskNumber))
        {
            ChangeOwnerStatusMessage = "Enter the ticket/task number authorizing this change.";
            return;
        }
        if (string.IsNullOrWhiteSpace(ChangeOwnerMailboxAddress))
        {
            ChangeOwnerStatusMessage = "Enter the shared mailbox's address.";
            return;
        }
        if (string.IsNullOrWhiteSpace(ChangeOwnerIdentities))
        {
            ChangeOwnerStatusMessage = "Enter at least one owner identity.";
            return;
        }

        var request = new ChangeSharedMailboxOwnerRequest
        {
            TaskNumber = ChangeOwnerTaskNumber.Trim(),
            MailboxAddress = ChangeOwnerMailboxAddress.Trim(),
            Mode = ChangeOwnerMode,
            OwnerIdentities = ChangeOwnerIdentities.Trim(),
        };

        IsChangingOwner = true;
        ChangeOwnerStatusMessage = "Applying owner change...";
        Transcript.BeginOperation($"Change owner: {ChangeOwnerMailboxAddress.Trim()} ({ChangeOwnerMode})");
        try
        {
            var result = await _sharedMailboxes.ChangeOwnerAsync(request);
            ChangeOwnerStatusMessage = result.Succeeded
                ? $"Owner(s) updated: {result.UpdatedOwners}"
                  + (result.WarningMessage is null ? "" : $" WARNING: {result.WarningMessage}")
                : $"Owner change failed: {result.ErrorMessage}";

            if (result.Succeeded)
            {
                ChangeOwnerIdentities = "";
                _ = RefreshOverviewAsync();
            }
        }
        catch (Exception ex)
        {
            ChangeOwnerStatusMessage = $"Owner change failed: {ex.Message}";
        }
        finally
        {
            IsChangingOwner = false;
        }
    }
}
