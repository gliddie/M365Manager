using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Controls;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Core.SharedMailboxes;
using M365Manager.Services;

namespace M365Manager.ViewModels;

public sealed partial class SharedMailboxesViewModel : ObservableObject, ITabbedPage
{
    /// <summary>Which tab the view shows; set by the shell when a Home tile navigates here.</summary>
    [ObservableProperty] private int _selectedTabIndex;

    private readonly ISharedMailboxService _sharedMailboxes;
    private readonly ISharedMailboxNamingService _naming;
    private readonly IDialogService _dialogs;

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
    [ObservableProperty] private string _statusMessage = "Fill in the fields and click Create shared mailbox.";

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

    // --- Rename ---
    //
    // Two-stage, like the legacy RenameShrMbx.ps1 dialog: look the mailbox up first, see what the
    // rename would touch (its access groups, its owners), then confirm. A rename hits the mailbox
    // and up to three groups, so "what am I about to change" is worth a round-trip.
    [ObservableProperty] private string _renameTaskNumber = "";
    [ObservableProperty] private string _renameMailboxIdentity = "";
    [ObservableProperty] private string _renameLocation = "";
    [ObservableProperty] private string _renameName = "";
    [ObservableProperty] private string _renameNewDisplayName = "";
    [ObservableProperty] private string _renameNewAddress = "";
    [ObservableProperty] private bool _isRenameNameEdited;
    [ObservableProperty] private bool _isRenameAddressEdited;
    [ObservableProperty] private bool _isRenaming;
    [ObservableProperty] private bool _isRenameLookupBusy;
    [ObservableProperty] private string _renameStatusMessage = "Look the mailbox up first, then enter its new name.";

    // Status severity, so success and failure stop looking identical. Set explicitly at each
    // outcome rather than inferred from the wording.
    [ObservableProperty] private AlertSeverity _statusSeverity = AlertSeverity.Info;
    [ObservableProperty] private AlertSeverity _changeOwnerStatusSeverity = AlertSeverity.Info;
    [ObservableProperty] private AlertSeverity _renameStatusSeverity = AlertSeverity.Info;
    [ObservableProperty] private AlertSeverity _overviewStatusSeverity = AlertSeverity.Info;

    private void Status(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        StatusSeverity = severity;
        StatusMessage = text;
    }

    private void ChangeOwnerStatus(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        ChangeOwnerStatusSeverity = severity;
        ChangeOwnerStatusMessage = text;
    }

    private void RenameStatus(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        RenameStatusSeverity = severity;
        RenameStatusMessage = text;
    }

    private void OverviewStatus(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        OverviewStatusSeverity = severity;
        OverviewStatusMessage = text;
    }

    /// <summary>What the looked-up mailbox is today - name, address, owners, access groups.</summary>
    [ObservableProperty] private string _renameCurrentSummary = "";

    /// <summary>What the access groups will be called after the rename. Empty until both halves are known.</summary>
    [ObservableProperty] private string _renameGroupSummary = "";

    /// <summary>Gates the Rename button: renaming a mailbox nobody looked up first is guesswork.</summary>
    [ObservableProperty] private bool _hasRenamePreview;

    private SharedMailboxRenamePreview? _renamePreview;
    private bool _applyingRenamePreview;
    private int _renamePreviewRequestVersion;

    /// <summary>
    /// True while the location/name fields still hold what a lookup put there. Looking up a
    /// different mailbox may then re-fill them; once the operator has typed, they are theirs.
    /// </summary>
    private bool _renameFieldsPrefilled;

    /// <summary>Guards <see cref="_renameFieldsPrefilled"/> while the lookup fills those fields in.</summary>
    private bool _applyingRenamePrefill;

    // --- Overview grid (read-only, straight from the SQL cache) ---
    [ObservableProperty] private string _overviewFilterText = "";
    [ObservableProperty] private bool _isOverviewBusy;
    [ObservableProperty] private string _overviewStatusMessage = "";

    /// <summary>
    /// Row picked in the overview grid. Selecting one fills the identity into the manage forms,
    /// which previously had to be typed into each of them by hand even though the grid listing the
    /// very same mailboxes sat right there. Nothing else about those forms changes.
    /// </summary>
    [ObservableProperty] private SharedMailboxOverviewRow? _selectedMailbox;

    partial void OnSelectedMailboxChanged(SharedMailboxOverviewRow? value)
    {
        if (value is null)
            return;

        // The address, not the display name: it is unique, whereas a display name can match more
        // than one mailbox and would send both forms into their "be more specific" error path.
        var identity = string.IsNullOrWhiteSpace(value.PrimarySmtpAddress)
            ? value.DisplayName
            : value.PrimarySmtpAddress;

        ChangeOwnerMailboxAddress = identity;
        RenameMailboxIdentity = identity;
    }

    public ObservableCollection<SharedMailboxOverviewRow> SharedMailboxes { get; } = new();

    /// <summary>Filtered view over <see cref="SharedMailboxes"/> that the grid binds to (and export reads from).</summary>
    public ICollectionView SharedMailboxesOverviewView { get; }

    /// <summary>Backs the PowerShell console on this page - see PowerShellConsole.</summary>
    public IPowerShellTranscript Transcript { get; }

    public SharedMailboxesViewModel(ISharedMailboxService sharedMailboxes, ISharedMailboxNamingService naming, ISettingsService settings, IPowerShellTranscript transcript, IDialogService dialogs)
    {
        _sharedMailboxes = sharedMailboxes;
        _naming = naming;
        _dialogs = dialogs;
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

            // An empty cache is not an achievement - the empty state next to it explains what to
            // do, and a green tick above that would just contradict it.
            OverviewStatus($"{SharedMailboxes.Count} shared mailbox(es) loaded.",
                SharedMailboxes.Count > 0 ? AlertSeverity.Success : AlertSeverity.Info);
        }
        catch (Exception ex)
        {
            OverviewStatus($"Could not load shared mailboxes: {ex.Message}", AlertSeverity.Error);
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
            OverviewStatus("Nothing to export - the (filtered) list is empty.", AlertSeverity.Error);
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
            OverviewStatus($"Exported {filtered.Count} shared mailbox(es) to {dialog.FileName}.", AlertSeverity.Success);
        }
        catch (Exception ex)
        {
            OverviewStatus($"Export failed: {ex.Message}", AlertSeverity.Error);
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
            Status(error, AlertSeverity.Error);
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
            StatusSeverity = result.Succeeded ? (result.WarningMessage is null ? AlertSeverity.Success : AlertSeverity.Warning) : AlertSeverity.Error;
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
            Status($"Shared mailbox creation failed: {ex.Message}", AlertSeverity.Error);
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
            ChangeOwnerStatus("Enter the ticket/task number authorizing this change.", AlertSeverity.Error);
            return;
        }
        if (string.IsNullOrWhiteSpace(ChangeOwnerMailboxAddress))
        {
            ChangeOwnerStatus("Enter the shared mailbox's display name or address.", AlertSeverity.Error);
            return;
        }
        if (string.IsNullOrWhiteSpace(ChangeOwnerIdentities))
        {
            ChangeOwnerStatus("Enter at least one owner identity.", AlertSeverity.Error);
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
            ChangeOwnerStatusSeverity = result.Succeeded ? (result.WarningMessage is null ? AlertSeverity.Success : AlertSeverity.Warning) : AlertSeverity.Error;
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
            ChangeOwnerStatus($"Owner change failed: {ex.Message}", AlertSeverity.Error);
        }
        finally
        {
            IsChangingOwner = false;
        }
    }

    // ----- Rename -----

    /// <summary>A preview belongs to one mailbox; typing a different one has to invalidate it.</summary>
    partial void OnRenameMailboxIdentityChanged(string value)
    {
        if (_renamePreview is null)
            return;

        _renamePreview = null;
        HasRenamePreview = false;
        RenameCurrentSummary = "";
        RenameGroupSummary = "";
        RenameStatusMessage = "Look the mailbox up again - the identity changed.";
    }

    partial void OnRenameLocationChanged(string value)
    {
        if (!_applyingRenamePrefill)
            _renameFieldsPrefilled = false;
        _ = RefreshRenamePreviewAsync();
    }

    partial void OnRenameNameChanged(string value)
    {
        if (!_applyingRenamePrefill)
            _renameFieldsPrefilled = false;
        _ = RefreshRenamePreviewAsync();
    }

    partial void OnRenameNewDisplayNameChanged(string value)
    {
        if (!_applyingRenamePreview)
            IsRenameNameEdited = true;
        UpdateRenameGroupSummary();
    }

    partial void OnRenameNewAddressChanged(string value)
    {
        if (!_applyingRenamePreview)
            IsRenameAddressEdited = true;
    }

    /// <summary>Reads the mailbox back from Exchange so the operator sees what a rename would touch.</summary>
    [RelayCommand]
    private async Task LookUpRenameMailboxAsync()
    {
        if (string.IsNullOrWhiteSpace(RenameMailboxIdentity))
        {
            RenameStatus("Enter the shared mailbox's display name or address.", AlertSeverity.Error);
            return;
        }

        IsRenameLookupBusy = true;
        RenameStatusMessage = "Looking the mailbox up...";
        Transcript.BeginOperation($"Look up shared mailbox: {RenameMailboxIdentity.Trim()}");
        try
        {
            var preview = await _sharedMailboxes.PreviewRenameAsync(RenameMailboxIdentity.Trim());
            _renamePreview = preview;
            HasRenamePreview = true;

            var groups = preview.AccessGroups.Count == 0
                ? "none found - only the mailbox itself would be renamed"
                : string.Join(", ", preview.AccessGroups.Select(g => $".{g.Key} {g.Value}"));

            RenameCurrentSummary =
                $"Name: {preview.DisplayName}\n"
                + $"Address: {preview.PrimarySmtpAddress}\n"
                + $"Alias: {preview.Alias}\n"
                + $"Owner(s): {(preview.Owners.Length == 0 ? "none resolved - no confirmation e-mail will be sent" : preview.Owners)}\n"
                + $"Access groups: {groups}";

            // The convention's own split: everything up to the first space is the location. Only
            // filled in while the fields are empty or still hold a previous lookup's values, so a
            // second lookup re-fills them but never overwrites something the operator typed.
            var untouched = (string.IsNullOrWhiteSpace(RenameLocation) && string.IsNullOrWhiteSpace(RenameName))
                            || _renameFieldsPrefilled;
            var space = preview.DisplayName.IndexOf(' ');
            if (untouched && space > 0)
            {
                _applyingRenamePrefill = true;
                try
                {
                    RenameLocation = preview.DisplayName[..space];
                    RenameName = preview.DisplayName[(space + 1)..];
                    _renameFieldsPrefilled = true;
                }
                finally
                {
                    _applyingRenamePrefill = false;
                }
            }

            await RefreshRenamePreviewAsync();
            UpdateRenameGroupSummary();
            RenameStatus("Mailbox found. Enter the new location/name, then confirm and rename.", AlertSeverity.Error);
        }
        catch (Exception ex)
        {
            _renamePreview = null;
            HasRenamePreview = false;
            RenameCurrentSummary = "";
            RenameGroupSummary = "";
            RenameStatus($"Lookup failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsRenameLookupBusy = false;
        }
    }

    /// <summary>Drops the manual corrections and goes back to what the naming convention produces.</summary>
    [RelayCommand]
    private void ResetRenamePreview()
    {
        IsRenameNameEdited = false;
        IsRenameAddressEdited = false;
        _ = RefreshRenamePreviewAsync();
    }

    private async Task RefreshRenamePreviewAsync()
    {
        var version = ++_renamePreviewRequestVersion;
        var location = RenameLocation;
        var name = RenameName;

        // The new address stays on the mailbox's current domain - a rename is about the name, and
        // the operator can still type a different address into the field by hand.
        var domain = _renamePreview?.Domain ?? "";

        if (string.IsNullOrWhiteSpace(location) || string.IsNullOrWhiteSpace(name))
        {
            ApplyRenamePreview(
                IsRenameNameEdited ? RenameNewDisplayName : "",
                IsRenameAddressEdited ? RenameNewAddress : "");
            return;
        }

        try
        {
            var (displayName, localPart) = await _naming.BuildNameAsync(location, name);
            if (version != _renamePreviewRequestVersion)
                return;

            ApplyRenamePreview(
                IsRenameNameEdited ? RenameNewDisplayName : displayName,
                IsRenameAddressEdited || domain.Length == 0
                    ? RenameNewAddress
                    : _naming.BuildAddress(localPart, domain));
        }
        catch
        {
            // Preview is best-effort (e.g. SQL not reachable yet) - leave the last known value.
        }
    }

    private void ApplyRenamePreview(string displayName, string address)
    {
        _applyingRenamePreview = true;
        try
        {
            RenameNewDisplayName = displayName;
            RenameNewAddress = address;
        }
        finally
        {
            _applyingRenamePreview = false;
        }
    }

    private void UpdateRenameGroupSummary()
    {
        var newName = RenameNewDisplayName.Trim();
        if (_renamePreview is null || newName.Length == 0 || _renamePreview.AccessGroups.Count == 0)
        {
            RenameGroupSummary = "";
            return;
        }

        RenameGroupSummary = string.Join("\n", _renamePreview.AccessGroups
            .Select(g => $".{g.Key}  {g.Value}  ->  {_naming.BuildAccessGroupName(newName, g.Key)}"));
    }

    [RelayCommand]
    private async Task RenameSharedMailboxAsync()
    {
        if (!ValidateRename(out var error))
        {
            RenameStatus(error, AlertSeverity.Error);
            return;
        }

        var identity = RenameMailboxIdentity.Trim();

        // A rename touches the mailbox and up to three access groups at once, and the old inline
        // checkbox could be ticked long before the button was pressed. The dialog asks at the
        // moment of the decision and lists exactly which groups are going with it.
        var groupCount = _renamePreview?.AccessGroups.Count ?? 0;
        var detail = groupCount == 0
            ? "This mailbox has no .ED/.AU/.RE access group, so only the mailbox itself is renamed."
            : $"These access groups are renamed with it:\n{RenameGroupSummary}";

        if (!_dialogs.ConfirmDestructive(
                $"Rename {_renamePreview?.DisplayName ?? identity}?",
                $"The mailbox becomes '{RenameNewDisplayName.Trim()}' ({RenameNewAddress.Trim()}). "
                + "The previous address is kept as an alias, so mail already addressed to it keeps arriving.",
                "Rename mailbox",
                detail))
            return;

        IsRenaming = true;
        RenameStatus("Starting shared mailbox rename...");
        Transcript.BeginOperation($"Rename shared mailbox: {identity} -> {RenameNewDisplayName.Trim()}");
        try
        {
            var result = await _sharedMailboxes.RenameAsync(new RenameSharedMailboxRequest
            {
                TaskNumber = RenameTaskNumber.Trim(),
                MailboxIdentity = identity,
                Location = RenameLocation.Trim(),
                Name = RenameName.Trim(),
                // Whatever is in the two fields wins - either the computed value or a correction.
                DisplayNameOverride = RenameNewDisplayName.Trim(),
                AddressOverride = RenameNewAddress.Trim(),
            }, msg => RenameStatusMessage = msg);

            if (result.Succeeded)
            {
                var groups = result.RenamedGroups.Count == 0
                    ? " No access group needed renaming."
                    : $" Renamed access group(s): {string.Join("; ", result.RenamedGroups)}.";

                // Only claimed when the address really changed - a name-only rename keeps no alias.
                var alias = string.Equals(result.PreviousPrimarySmtpAddress, result.PrimarySmtpAddress, StringComparison.OrdinalIgnoreCase)
                    ? ""
                    : $" The previous address {result.PreviousPrimarySmtpAddress} is kept as an alias.";

                RenameStatusMessage =
                    $"'{result.PreviousDisplayName}' renamed to '{result.DisplayName}' ({result.PrimarySmtpAddress})."
                    + alias
                    + groups
                    + (result.WarningMessage is null ? "" : $" WARNING: {result.WarningMessage}");

                ClearRenameForm();
                _ = RefreshOverviewAsync();
            }
            else
            {
                RenameStatus($"Rename failed: {result.ErrorMessage}", AlertSeverity.Error);
            }
        }
        catch (Exception ex)
        {
            RenameStatus($"Rename failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsRenaming = false;
        }
    }

    private bool ValidateRename(out string error)
    {
        if (string.IsNullOrWhiteSpace(RenameTaskNumber)) { error = "Enter the ticket/task number authorizing this change."; return false; }
        if (string.IsNullOrWhiteSpace(RenameMailboxIdentity)) { error = "Enter the shared mailbox's display name or address."; return false; }
        if (!HasRenamePreview) { error = "Look the mailbox up first - the rename needs to know which access groups it has."; return false; }
        if (string.IsNullOrWhiteSpace(RenameNewDisplayName)) { error = "The new display name is empty."; return false; }
        if (string.IsNullOrWhiteSpace(RenameNewAddress) || !RenameNewAddress.Contains('@')) { error = "The new address is not a valid e-mail address."; return false; }
        error = "";
        return true;
    }

    private void ClearRenameForm()
    {
        RenameTaskNumber = "";
        RenameLocation = "";
        RenameName = "";
        IsRenameNameEdited = false;
        IsRenameAddressEdited = false;
        _renameFieldsPrefilled = false;
        ApplyRenamePreview("", "");
        RenameCurrentSummary = "";
        RenameGroupSummary = "";
        HasRenamePreview = false;
        // Cleared before the identity on purpose: OnRenameMailboxIdentityChanged bails out when
        // there is no preview left, so the success message just written survives this reset.
        _renamePreview = null;
        RenameMailboxIdentity = "";
    }
}
