using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Controls;
using M365Manager.Core.Exchange;
using M365Manager.Core.Groups;
using M365Manager.Core.M365;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Data.Logging;
using M365Manager.Services;

namespace M365Manager.ViewModels;

public sealed partial class GroupsViewModel : ObservableObject, ITabbedPage
{
    /// <summary>Which tab the view shows; set by the shell when a Home tile navigates here.</summary>
    [ObservableProperty] private int _selectedTabIndex;

    private readonly IExchangeService _exchange;
    private readonly IGroupAdminService _groupAdmin;
    private readonly IGroupNamingService _naming;
    private readonly ISettingsService _settings;
    private readonly IDialogService _dialogs;
    private readonly ILogService _log;
    private readonly IM365AuthService _auth;

    [ObservableProperty] private string _searchText = "";
    [ObservableProperty] private bool _isConnected;
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = "Connect to Exchange Online to search groups.";

    [ObservableProperty] private DistributionGroupInfo? _selectedGroup;
    [ObservableProperty] private string _ownersText = "";

    /// <summary>Device-code sign-in instructions, shown prominently while connecting.</summary>
    [ObservableProperty] private string _deviceCodeMessage = "";

    /// <summary>Just the sign-in code, for the copy button.</summary>
    [ObservableProperty] private string _deviceCode = "";

    [ObservableProperty] private GroupMemberInfo? _selectedMember;
    [ObservableProperty] private string _newMemberAddress = "";

    /// <summary>ServiceNow task number required before any add/remove change. Not auto-cleared.</summary>
    [ObservableProperty] private string _taskNumber = "";

    /// <summary>Synced from the Members DataGrid's multi-selection (see GroupsView.xaml.cs).</summary>
    private IReadOnlyList<GroupMemberInfo> _selectedMembers = Array.Empty<GroupMemberInfo>();

    /// <summary>Free-text filter applied to <see cref="MembersView"/> (name, address, UPN, type, path).</summary>
    [ObservableProperty] private string _memberFilterText = "";

    private bool _browserOpened;

    public ObservableCollection<DistributionGroupInfo> SearchResults { get; } = new();
    public ObservableCollection<GroupMemberInfo> Members { get; } = new();

    /// <summary>Filtered view over <see cref="Members"/> that the DataGrid binds to.</summary>
    public ICollectionView MembersView { get; }

    /// <summary>Backs the PowerShell console on this page - see PowerShellConsole.</summary>
    public IPowerShellTranscript Transcript { get; }

    // ================= Create =================

    [ObservableProperty] private string _createTaskNumber = "";

    /// <summary>Who gets the confirmation e-mail. Empty means none is sent - see GroupCreationRequest.</summary>
    [ObservableProperty] private string _createRequesterIdentity = "";
    [ObservableProperty] private string _createRawName = "";
    [ObservableProperty] private GroupKind _createKind = GroupKind.Distribution;
    [ObservableProperty] private bool _createIsDynamic;
    [ObservableProperty] private string _createDomain = "";
    [ObservableProperty] private bool _createIsTemporary;
    [ObservableProperty] private DateTime? _createRemoveOn = DateTime.Today.AddMonths(3);
    [ObservableProperty] private string _createRemoveCharacter = "";
    [ObservableProperty] private string _createOwners = "";
    [ObservableProperty] private string _createMembers = "";
    [ObservableProperty] private string _createAuthorizedSenders = "";
    [ObservableProperty] private string _createAdditionalAliases = "";
    [ObservableProperty] private bool _createAllowExternalSenders;
    [ObservableProperty] private bool _isCreateBusy;
    [ObservableProperty] private string _createStatusMessage = "Fill in the fields and click Create.";

    // Editable preview, same mechanism as the Shared Mailboxes form.
    [ObservableProperty] private string _previewName = "";
    [ObservableProperty] private string _previewAddress = "";
    [ObservableProperty] private bool _isPreviewNameEdited;
    [ObservableProperty] private bool _isPreviewAddressEdited;
    [ObservableProperty] private string _previewProblems = "";
    private bool _applyingPreview;
    private int _previewRequestVersion;

    // Dynamic group conditions, pre-filled with the values from the legacy script.
    [ObservableProperty] private string _dynamicIncludedRecipients = "MailboxUsers";
    [ObservableProperty] private string _dynamicCompany = "UL.COM";
    [ObservableProperty] private string _dynamicCustomAttribute1 = "Employee";
    [ObservableProperty] private string _dynamicCustomAttribute2 = "A";
    [ObservableProperty] private string _dynamicCustomAttribute3 = "";
    [ObservableProperty] private string _dynamicCustomAttribute8 = "";
    [ObservableProperty] private string _dynamicNotes = "";
    [ObservableProperty] private string _dynamicMailTip = "";

    public IReadOnlyList<GroupKind> GroupKinds { get; } = Enum.GetValues<GroupKind>();
    public IReadOnlyList<ListChangeMode> ChangeModes { get; } = Enum.GetValues<ListChangeMode>();

    public bool IsStaticCreate => !CreateIsDynamic;

    // ================= Manage selected =================

    [ObservableProperty] private string _manageTaskNumber = "";

    /// <summary>Requester for the manage panel. Alias changes send no mail - they have no template.</summary>
    [ObservableProperty] private string _manageRequesterIdentity = "";

    [ObservableProperty] private bool _isManageBusy;
    [ObservableProperty] private string _manageStatusMessage = "";

    [ObservableProperty] private ListChangeMode _ownerChangeMode = ListChangeMode.Add;
    [ObservableProperty] private string _ownerIdentities = "";

    [ObservableProperty] private ListChangeMode _senderChangeMode = ListChangeMode.Add;
    [ObservableProperty] private string _senderIdentities = "";
    [ObservableProperty] private string _currentAuthorizedSenders = "";

    [ObservableProperty] private ListChangeMode _aliasChangeMode = ListChangeMode.Add;
    [ObservableProperty] private string _aliasIdentities = "";
    [ObservableProperty] private string _currentAliases = "";

    [ObservableProperty] private string _replaceMembersText = "";

    [ObservableProperty] private string _renameNewName = "";
    [ObservableProperty] private string _renameNewAddress = "";

    // ================= Status severity =================
    //
    // The three status lines used to render as identical grey text, so a failure looked exactly
    // like a success. Severity is set explicitly at each outcome rather than guessed from the
    // wording - "Removed 0 members" is a success, "Could not find" is not, and no substring rule
    // gets that right for long.

    [ObservableProperty] private AlertSeverity _statusSeverity = AlertSeverity.Info;
    [ObservableProperty] private AlertSeverity _createStatusSeverity = AlertSeverity.Info;
    [ObservableProperty] private AlertSeverity _manageStatusSeverity = AlertSeverity.Info;

    private void Status(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        StatusSeverity = severity;
        StatusMessage = text;
    }

    private void CreateStatus(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        CreateStatusSeverity = severity;
        CreateStatusMessage = text;
    }

    private void ManageStatus(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        ManageStatusSeverity = severity;
        ManageStatusMessage = text;
    }


    public GroupsViewModel(
        IExchangeService exchange,
        ILogService log,
        IM365AuthService auth,
        IPowerShellTranscript transcript,
        IGroupAdminService groupAdmin,
        IGroupNamingService naming,
        ISettingsService settings,
        IDialogService dialogs)
    {
        _exchange = exchange;
        _log = log;
        _auth = auth;
        _groupAdmin = groupAdmin;
        _naming = naming;
        _settings = settings;
        _dialogs = dialogs;
        Transcript = transcript;

        MembersView = CollectionViewSource.GetDefaultView(Members);
        MembersView.Filter = FilterMember;

        CreateDomain = settings.Current.M365.DefaultMailDomain;
    }

    partial void OnMemberFilterTextChanged(string value) => MembersView.Refresh();

    private bool FilterMember(object obj)
    {
        if (string.IsNullOrWhiteSpace(MemberFilterText))
            return true;
        if (obj is not GroupMemberInfo m)
            return true;

        var term = MemberFilterText.Trim();
        return Contains(m.DisplayName, term)
            || Contains(m.PrimarySmtpAddress, term)
            || Contains(m.UserPrincipalName, term)
            || Contains(m.RecipientType, term)
            || Contains(m.Path, term);
    }

    private static bool Contains(string? value, string term)
        => !string.IsNullOrEmpty(value) && value.Contains(term, StringComparison.OrdinalIgnoreCase);

    /// <summary>Called from GroupsView code-behind when the Members DataGrid's selection changes.</summary>
    public void UpdateMemberSelection(IEnumerable<GroupMemberInfo> items) => _selectedMembers = items.ToList();

    /// <summary>
    /// Reflects the shared ExchangeService's connection state onto this page. Called by
    /// MainViewModel after the app-wide startup connect finishes, so this page shows
    /// "Connected" without the user having to click this page's own Connect button.
    /// </summary>
    public void RefreshConnectionState()
    {
        IsConnected = _exchange.IsConnected;
        if (IsConnected)
            Status("Connected to Exchange Online. Enter a name and search.", AlertSeverity.Error);
    }

    [RelayCommand]
    private async Task ConnectAsync()
    {
        IsBusy = true;
        _browserOpened = false;
        DeviceCodeMessage = "";
        DeviceCode = "";
        StatusMessage = "Connecting to Exchange Online...";
        var dispatcher = System.Windows.Application.Current?.Dispatcher;
        try
        {
            await _exchange.ConnectAsync(prompt =>
            {
                void Show()
                {
                    // Accumulate all sign-in output so the full instructions stay visible.
                    DeviceCodeMessage = string.IsNullOrEmpty(DeviceCodeMessage) ? prompt : $"{DeviceCodeMessage}\n{prompt}";
                    StatusMessage = prompt;
                    ExtractDeviceCode(prompt);
                    TryOpenBrowser(prompt);
                }

                if (dispatcher is not null)
                    dispatcher.Invoke(Show);
                else
                    Show();
            });

            IsConnected = _exchange.IsConnected;
            DeviceCodeMessage = "";
            DeviceCode = "";
            Status("Connected to Exchange Online. Enter a name and search.", AlertSeverity.Error);
            await SafeLogAsync("Connect", null, Severity.Success, "Connected to Exchange Online.");
        }
        catch (Exception ex)
        {
            DeviceCodeMessage = "";
            DeviceCode = "";
            Status($"Connection failed: {ex.Message}", AlertSeverity.Error);
            await SafeLogAsync("Connect", null, Severity.Error, $"Exchange connect failed: {ex.Message}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    private void ExtractDeviceCode(string prompt)
    {
        var code = DeviceCodePrompt.ExtractCode(prompt);
        if (code is not null)
            DeviceCode = code;
    }

    [RelayCommand]
    private void CopyCode()
    {
        if (string.IsNullOrEmpty(DeviceCode))
            return;
        try
        {
            System.Windows.Clipboard.SetText(DeviceCode);
            StatusMessage = $"Code {DeviceCode} copied to clipboard.";
        }
        catch
        {
            // Clipboard can occasionally be locked by another app; ignore.
        }
    }

    private void TryOpenBrowser(string prompt)
    {
        if (_browserOpened)
            return;

        _browserOpened = DeviceCodePrompt.TryOpenBrowser(prompt);
    }

    [RelayCommand]
    private async Task SearchAsync()
    {
        if (!IsConnected)
        {
            Status("Please connect to Exchange Online first.", AlertSeverity.Error);
            return;
        }
        if (string.IsNullOrWhiteSpace(SearchText))
        {
            Status("Enter a group name, alias or address to search.", AlertSeverity.Error);
            return;
        }

        SearchResults.Clear();
        Members.Clear();
        SelectedGroup = null;
        SelectedMember = null;
        OwnersText = "";
        NewMemberAddress = "";
        MemberFilterText = "";

        IsBusy = true;
        StatusMessage = $"Searching for '{SearchText}'...";
        try
        {
            var groups = await _exchange.SearchGroupsAsync(SearchText);
            foreach (var g in groups)
                SearchResults.Add(g);

            Status($"{SearchResults.Count} group(s) found.", AlertSeverity.Success);
            await SafeLogAsync("Search", SearchText, Severity.Info, $"Group search '{SearchText}' returned {SearchResults.Count}.");
        }
        catch (Exception ex)
        {
            Status($"Search failed: {ex.Message}", AlertSeverity.Error);
            await SafeLogAsync("Search", SearchText, Severity.Error, $"Group search failed: {ex.Message}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void ClearResults()
    {
        SearchText = "";
        SearchResults.Clear();
        Members.Clear();
        SelectedGroup = null;
        SelectedMember = null;
        OwnersText = "";
        NewMemberAddress = "";
        MemberFilterText = "";
        _selectedMembers = Array.Empty<GroupMemberInfo>();
        StatusMessage = "Search cleared.";
    }

    partial void OnSelectedGroupChanged(DistributionGroupInfo? value)
    {
        OwnersText = value is null ? "" : string.Join(", ", value.ManagedBy);
        ManageStatusMessage = "";
        _ = LoadMembersAsync(value);
        _ = LoadGroupDetailsAsync(value);
    }

    private async Task LoadMembersAsync(DistributionGroupInfo? group)
    {
        Members.Clear();
        MemberFilterText = "";
        _selectedMembers = Array.Empty<GroupMemberInfo>();
        if (group is null)
            return;

        IsBusy = true;
        StatusMessage = $"Loading members of {group.DisplayName}...";
        try
        {
            var members = await _exchange.GetMembersAsync(group);
            foreach (var m in members)
                Members.Add(m);

            Status($"{group.DisplayName}: {Members.Count} member(s).", AlertSeverity.Success);
            await SafeLogAsync("ViewMembers", group.PrimarySmtpAddress, Severity.Info,
                $"Viewed {Members.Count} members of {group.DisplayName}.");
        }
        catch (Exception ex)
        {
            Status($"Could not load members: {ex.Message}", AlertSeverity.Error);
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task AddMemberAsync()
    {
        if (SelectedGroup is null)
        {
            Status("Select a group first.", AlertSeverity.Error);
            return;
        }

        var identities = NewMemberAddress
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();
        if (identities.Count == 0)
        {
            Status("Enter a SamAccountName, email or UPN to add (comma-separate for multiple).", AlertSeverity.Error);
            return;
        }
        if (!TryRequireTaskNumber(out var task))
            return;

        var group = SelectedGroup;
        IsBusy = true;
        Transcript.BeginOperation($"Add member(s) to {group.DisplayName}");
        StatusMessage = identities.Count == 1
            ? $"Adding {identities[0]} to {group.DisplayName}..."
            : $"Adding {identities.Count} members to {group.DisplayName}...";
        try
        {
            var results = await _exchange.AddMembersAsync(group, identities);
            foreach (var r in results)
            {
                await SafeLogAsync("AddMember", group.PrimarySmtpAddress,
                    r.Succeeded ? Severity.Success : Severity.Error,
                    r.Succeeded
                        ? $"Added {r.Identity} to {group.DisplayName}."
                        : $"Add {r.Identity} to {group.DisplayName} failed: {r.Error}",
                    task);
            }

            NewMemberAddress = "";
            await LoadMembersAsync(group);
            Status(SummarizeOutcome(results, "Added", group.DisplayName, "to"), AlertSeverity.Success);
        }
        catch (Exception ex)
        {
            Status($"Add failed: {ex.Message}", AlertSeverity.Error);
            await SafeLogAsync("AddMember", group.PrimarySmtpAddress, Severity.Error, $"Add to {group.DisplayName} failed: {ex.Message}", task);
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task RemoveMemberAsync()
    {
        if (SelectedGroup is null)
        {
            Status("Select a group first.", AlertSeverity.Error);
            return;
        }

        var members = _selectedMembers.Count > 0
            ? _selectedMembers
            : SelectedMember is not null ? new[] { SelectedMember } : Array.Empty<GroupMemberInfo>();
        if (members.Count == 0)
        {
            Status("Select one or more members to remove (click a row, or Ctrl/Shift-click for multiple).", AlertSeverity.Error);
            return;
        }
        if (!TryRequireTaskNumber(out var task))
            return;

        var names = members.Select(m => m.DisplayName).ToList();
        var namesText = names.Count <= 10 ? string.Join(", ", names) : $"{string.Join(", ", names.Take(10))}, and {names.Count - 10} more";

        if (!_dialogs.ConfirmDestructive(
                $"Remove {members.Count} member(s) from {SelectedGroup.DisplayName}?",
                "They lose access to anything delivered through this group. You can add them back afterwards.",
                members.Count == 1 ? "Remove member" : $"Remove {members.Count} members",
                namesText))
            return;

        var group = SelectedGroup;
        var identities = members.Select(m => string.IsNullOrWhiteSpace(m.PrimarySmtpAddress) ? m.DisplayName : m.PrimarySmtpAddress).ToList();
        IsBusy = true;
        Transcript.BeginOperation($"Remove member(s) from {group.DisplayName}");
        StatusMessage = $"Removing {members.Count} member(s) from {group.DisplayName}...";
        try
        {
            var results = await _exchange.RemoveMembersAsync(group, identities);
            foreach (var r in results)
            {
                await SafeLogAsync("RemoveMember", group.PrimarySmtpAddress,
                    r.Succeeded ? Severity.Success : Severity.Error,
                    r.Succeeded
                        ? $"Removed {r.Identity} from {group.DisplayName}."
                        : $"Remove {r.Identity} from {group.DisplayName} failed: {r.Error}",
                    task);
            }

            await LoadMembersAsync(group);
            Status(SummarizeOutcome(results, "Removed", group.DisplayName, "from"), AlertSeverity.Success);
        }
        catch (Exception ex)
        {
            Status($"Remove failed: {ex.Message}", AlertSeverity.Error);
            await SafeLogAsync("RemoveMember", group.PrimarySmtpAddress, Severity.Error, $"Remove from {group.DisplayName} failed: {ex.Message}", task);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private bool TryRequireTaskNumber(out string task)
    {
        task = TaskNumber.Trim();
        if (task.Length > 0)
            return true;

        Status("Enter the ServiceNow task number authorizing this change before continuing.", AlertSeverity.Error);
        return false;
    }

    private static string SummarizeOutcome(IReadOnlyList<MemberOperationResult> results, string verbPast, string groupName, string preposition)
    {
        var succeeded = results.Count(r => r.Succeeded);
        var failed = results.Count - succeeded;
        return failed == 0
            ? $"{verbPast} {succeeded} member(s) {preposition} {groupName}."
            : $"{verbPast} {succeeded} of {results.Count} member(s) {preposition} {groupName}. {failed} failed - see Logs for details.";
    }

    [RelayCommand]
    private async Task SyncMembersAsync()
    {
        if (SelectedGroup is null)
        {
            Status("Select a group first.", AlertSeverity.Error);
            return;
        }

        var group = SelectedGroup;
        IsBusy = true;
        Transcript.BeginOperation($"Sync members of {group.DisplayName}");
        StatusMessage = $"Syncing members of {group.DisplayName} from Exchange...";
        try
        {
            var members = await _exchange.SyncMembersAsync(group);
            Members.Clear();
            foreach (var m in members)
                Members.Add(m);

            Status($"{group.DisplayName}: {Members.Count} direct member(s) synced.", AlertSeverity.Success);
            await SafeLogAsync("SyncMembers", group.PrimarySmtpAddress, Severity.Success,
                $"Synced {Members.Count} direct members of {group.DisplayName} from Exchange.");
        }
        catch (Exception ex)
        {
            Status($"Sync failed: {ex.Message}", AlertSeverity.Error);
            await SafeLogAsync("SyncMembers", group.PrimarySmtpAddress, Severity.Error,
                $"Sync members of {group.DisplayName} failed: {ex.Message}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void ExportToExcel()
    {
        if (SelectedGroup is null)
        {
            Status("Select a group first.", AlertSeverity.Error);
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = $"{SanitizeFileName(SelectedGroup.DisplayName)}.xlsx",
        };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            GroupExportService.ExportToExcel(dialog.FileName, SelectedGroup, Members.ToList());
            Status($"Exported to {dialog.FileName}.", AlertSeverity.Success);
        }
        catch (Exception ex)
        {
            Status($"Export failed: {ex.Message}", AlertSeverity.Error);
        }
    }

    [RelayCommand]
    private void ExportToCsv()
    {
        if (SelectedGroup is null)
        {
            Status("Select a group first.", AlertSeverity.Error);
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "CSV file (*.csv)|*.csv",
            FileName = $"{SanitizeFileName(SelectedGroup.DisplayName)}.csv",
        };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            GroupExportService.ExportToCsv(dialog.FileName, SelectedGroup, Members.ToList());
            Status($"Exported to {dialog.FileName}.", AlertSeverity.Success);
        }
        catch (Exception ex)
        {
            Status($"Export failed: {ex.Message}", AlertSeverity.Error);
        }
    }

    private static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return string.IsNullOrWhiteSpace(name) ? "Group" : name;
    }

    private async Task SafeLogAsync(string action, string? target, Severity severity, string message, string? taskNumber = null)
    {
        try
        {
            await _log.WriteAsync(new LogEntry
            {
                Area = "Groups",
                Action = action,
                TargetObject = target,
                TaskNumber = taskNumber,
                UserUpn = _auth.CurrentUser?.Upn,
                Severity = severity,
                EventCode = severity == Severity.Error ? "ERR" : "INFO",
                Message = message,
            });
        }
        catch
        {
            // Logging is best-effort; never let it break the UI action.
        }
    }

    // ================================================================
    // Create
    // ================================================================

    partial void OnCreateRawNameChanged(string value) => _ = RefreshPreviewAsync();
    partial void OnCreateDomainChanged(string value) => _ = RefreshPreviewAsync();
    partial void OnCreateIsTemporaryChanged(bool value) => _ = RefreshPreviewAsync();
    partial void OnCreateRemoveCharacterChanged(string value) => _ = RefreshPreviewAsync();

    partial void OnCreateIsDynamicChanged(bool value) => OnPropertyChanged(nameof(IsStaticCreate));

    // Once a preview field is edited by hand it stops being overwritten - typing on in the name
    // field must not silently discard the correction.
    partial void OnPreviewNameChanged(string value)
    {
        if (!_applyingPreview) IsPreviewNameEdited = true;
    }

    partial void OnPreviewAddressChanged(string value)
    {
        if (!_applyingPreview) IsPreviewAddressEdited = true;
    }

    [RelayCommand]
    private void ResetPreview()
    {
        IsPreviewNameEdited = false;
        IsPreviewAddressEdited = false;
        _ = RefreshPreviewAsync();
    }

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

    private async Task RefreshPreviewAsync()
    {
        var version = ++_previewRequestVersion;

        if (string.IsNullOrWhiteSpace(CreateRawName))
        {
            ApplyPreview(IsPreviewNameEdited ? PreviewName : "", IsPreviewAddressEdited ? PreviewAddress : "");
            PreviewProblems = "";
            return;
        }

        try
        {
            var names = await _naming.BuildAsync(CreateRawName, CreateIsTemporary, CreateRemoveCharacter, CreateDomain);
            if (version != _previewRequestVersion)
                return;

            ApplyPreview(
                IsPreviewNameEdited ? PreviewName : names.DisplayName,
                IsPreviewAddressEdited ? PreviewAddress : names.Address);

            PreviewProblems = PreviewName.Length > 64
                ? $"The name is {PreviewName.Length} characters - Exchange allows at most 64."
                : "";
        }
        catch
        {
            // Preview is best-effort (SQL may not be reachable) - keep the last value.
        }
    }

    [RelayCommand]
    private async Task CreateGroupAsync()
    {
        if (string.IsNullOrWhiteSpace(CreateTaskNumber))
        {
            CreateStatus("Enter the ticket/task number authorizing this creation.", AlertSeverity.Error);
            return;
        }

        IsCreateBusy = true;
        try
        {
            if (CreateIsDynamic)
                await CreateDynamicGroupAsync();
            else
                await CreateStaticGroupAsync();
        }
        catch (Exception ex)
        {
            CreateStatus($"Creation failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsCreateBusy = false;
        }
    }

    private async Task CreateStaticGroupAsync()
    {
        var request = new GroupCreationRequest
        {
            TaskNumber = CreateTaskNumber.Trim(),
            RequesterIdentity = CreateRequesterIdentity.Trim(),
            RawName = CreateRawName.Trim(),
            Kind = CreateKind,
            Domain = CreateDomain.Trim(),
            IsTemporary = CreateIsTemporary,
            RemoveOn = CreateIsTemporary ? CreateRemoveOn : null,
            RemoveCharacter = CreateRemoveCharacter,
            Owners = CreateOwners.Trim(),
            Members = CreateMembers.Trim(),
            AuthorizedSenders = CreateAuthorizedSenders.Trim(),
            AdditionalAliases = CreateAdditionalAliases.Trim(),
            AllowExternalSenders = CreateAllowExternalSenders,
            DisplayNameOverride = PreviewName.Trim(),
            AddressOverride = PreviewAddress.Trim(),
        };

        CreateStatusMessage = "Starting group creation...";
        Transcript.BeginOperation($"Create group: {PreviewName.Trim()}");

        var result = await _groupAdmin.CreateAsync(request, msg => CreateStatusMessage = msg);
        if (result.Succeeded)
        {
            CreateStatusSeverity = result.Warning is { Length: > 0 } ? AlertSeverity.Warning : AlertSeverity.Success;
            CreateStatusMessage = result.Warning is { Length: > 0 }
                ? $"'{result.DisplayName}' created - but note: {result.Warning}"
                : $"'{result.DisplayName}' ({result.Address}) created.";
            ClearCreateForm();
        }
        else
        {
            CreateStatus($"Creation failed: {result.ErrorMessage}", AlertSeverity.Error);
        }
    }

    private async Task CreateDynamicGroupAsync()
    {
        var request = new DynamicGroupCreationRequest
        {
            TaskNumber = CreateTaskNumber.Trim(),
            RequesterIdentity = CreateRequesterIdentity.Trim(),
            DisplayName = string.IsNullOrWhiteSpace(PreviewName) ? CreateRawName.Trim() : PreviewName.Trim(),
            Domain = CreateDomain.Trim(),
            AddressOverride = PreviewAddress.Trim(),
            IncludedRecipients = DynamicIncludedRecipients.Trim(),
            ConditionalCompany = DynamicCompany.Trim(),
            ConditionalCustomAttribute1 = DynamicCustomAttribute1.Trim(),
            ConditionalCustomAttribute2 = DynamicCustomAttribute2.Trim(),
            ConditionalCustomAttribute3 = DynamicCustomAttribute3.Trim(),
            ConditionalCustomAttribute8 = DynamicCustomAttribute8.Trim(),
            Notes = DynamicNotes.Trim(),
            MailTip = DynamicMailTip.Trim(),
        };

        CreateStatusMessage = "Starting dynamic group creation...";
        Transcript.BeginOperation($"Create dynamic group: {request.DisplayName}");

        var result = await _groupAdmin.CreateDynamicAsync(request, msg => CreateStatusMessage = msg);
        CreateStatusSeverity = result.Succeeded ? AlertSeverity.Success : AlertSeverity.Error;
        CreateStatusMessage = result.Succeeded
            ? $"'{result.DisplayName}' ({result.Address}) created."
              + (result.Warning is { Length: > 0 } w ? $" WARNING: {w}" : "")
            : $"Creation failed: {result.ErrorMessage}";

        if (result.Succeeded)
            ClearCreateForm();
    }

    private void ClearCreateForm()
    {
        CreateTaskNumber = "";
        CreateRequesterIdentity = "";
        CreateRawName = "";
        CreateOwners = "";
        CreateMembers = "";
        CreateAuthorizedSenders = "";
        CreateAdditionalAliases = "";
        CreateAllowExternalSenders = false;
        CreateIsTemporary = false;
        CreateRemoveCharacter = "";
        IsPreviewNameEdited = false;
        IsPreviewAddressEdited = false;
        ApplyPreview("", "");
        PreviewProblems = "";
    }

    // ================================================================
    // Manage selected group
    // ================================================================

    /// <summary>Loads the settings shown in the manage panel for the currently selected group.</summary>
    private async Task LoadGroupDetailsAsync(DistributionGroupInfo? group)
    {
        CurrentAuthorizedSenders = "";
        CurrentAliases = "";
        RenameNewName = "";
        RenameNewAddress = "";

        if (group is null)
            return;

        try
        {
            var identity = GroupIdentity(group);
            var details = await _groupAdmin.GetDetailsAsync(identity);
            if (details is null)
                return;

            CurrentAuthorizedSenders = details.AuthorizedSenders.Count == 0
                ? "(anyone may send)"
                : string.Join("\n", details.AuthorizedSenders);

            // EmailAddresses carries the smtp:/SMTP: prefixes; strip them for display.
            CurrentAliases = string.Join("\n", details.Aliases
                .Select(a => a.Contains(':') ? a[(a.IndexOf(':') + 1)..] : a));

            RenameNewName = details.DisplayName;
            RenameNewAddress = details.PrimarySmtpAddress;
        }
        catch (Exception ex)
        {
            ManageStatus($"Could not read group details: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
    }

    /// <summary>Prefers the SMTP address; a cached row stores the Graph id in Name, which Exchange won't resolve.</summary>
    private static string GroupIdentity(DistributionGroupInfo group)
        => string.IsNullOrWhiteSpace(group.PrimarySmtpAddress) ? group.Name : group.PrimarySmtpAddress;

    private bool ValidateManage(out string error)
    {
        if (SelectedGroup is null) { error = "Select a group first."; return false; }
        if (string.IsNullOrWhiteSpace(ManageTaskNumber)) { error = "Enter the ticket/task number authorizing this change."; return false; }
        error = "";
        return true;
    }

    private GroupListChangeRequest BuildListRequest(ListChangeMode mode, string identities) => new()
    {
        TaskNumber = ManageTaskNumber.Trim(),
        RequesterIdentity = ManageRequesterIdentity.Trim(),
        GroupIdentity = GroupIdentity(SelectedGroup!),
        Mode = mode,
        Identities = identities.Trim(),
        IsUnifiedGroup = SelectedGroup!.IsUnifiedGroup,
    };

    private void ReportManage(GroupOperationResult result, string success)
    {
        if (!result.Succeeded)
        {
            ManageStatus($"Failed: {result.ErrorMessage}", AlertSeverity.Error);
            return;
        }

        var failed = result.Results.Where(r => !r.Succeeded).ToList();
        ManageStatusSeverity = failed.Count == 0 ? AlertSeverity.Success : AlertSeverity.Error;
        ManageStatusMessage = failed.Count == 0
            ? success
            : $"{success} {failed.Count} failed: {string.Join("; ", failed.Select(f => $"{f.Identity} ({f.Error})"))}";
    }

    [RelayCommand]
    private async Task ChangeOwnersAsync()
    {
        if (!ValidateManage(out var error)) { ManageStatusMessage = error; return; }
        if (string.IsNullOrWhiteSpace(OwnerIdentities)) { ManageStatusMessage = "Enter at least one owner."; return; }

        IsManageBusy = true;
        ManageStatusMessage = "Applying owner change...";
        Transcript.BeginOperation($"{OwnerChangeMode} owner(s) on {SelectedGroup!.DisplayName}");
        try
        {
            var result = await _groupAdmin.ChangeOwnersAsync(BuildListRequest(OwnerChangeMode, OwnerIdentities));
            ReportManage(result, $"Owners updated ({OwnerChangeMode}).");
            if (result.Succeeded)
            {
                OwnerIdentities = "";
                await RefreshSelectedGroupAsync();
            }
        }
        catch (Exception ex)
        {
            ManageStatus($"Owner change failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    [RelayCommand]
    private async Task ChangeAuthorizedSendersAsync()
    {
        if (!ValidateManage(out var error)) { ManageStatusMessage = error; return; }
        if (string.IsNullOrWhiteSpace(SenderIdentities)) { ManageStatusMessage = "Enter at least one identity."; return; }

        IsManageBusy = true;
        ManageStatusMessage = "Applying authorized sender change...";
        Transcript.BeginOperation($"{SenderChangeMode} authorized sender(s) on {SelectedGroup!.DisplayName}");
        try
        {
            var result = await _groupAdmin.ChangeAuthorizedSendersAsync(BuildListRequest(SenderChangeMode, SenderIdentities));
            ReportManage(result, $"Authorized senders updated ({SenderChangeMode}).");
            if (result.Succeeded)
            {
                SenderIdentities = "";
                await LoadGroupDetailsAsync(SelectedGroup);
            }
        }
        catch (Exception ex)
        {
            ManageStatus($"Authorized sender change failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    [RelayCommand]
    private async Task ChangeAliasesAsync()
    {
        if (!ValidateManage(out var error)) { ManageStatusMessage = error; return; }
        if (string.IsNullOrWhiteSpace(AliasIdentities)) { ManageStatusMessage = "Enter at least one address."; return; }

        IsManageBusy = true;
        ManageStatusMessage = "Applying alias change...";
        Transcript.BeginOperation($"{AliasChangeMode} alias(es) on {SelectedGroup!.DisplayName}");
        try
        {
            var result = await _groupAdmin.ChangeAliasesAsync(BuildListRequest(AliasChangeMode, AliasIdentities));
            ReportManage(result, $"Aliases updated ({AliasChangeMode}).");
            if (result.Succeeded)
            {
                AliasIdentities = "";
                await LoadGroupDetailsAsync(SelectedGroup);
            }
        }
        catch (Exception ex)
        {
            ManageStatus($"Alias change failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    [RelayCommand]
    private async Task ReplaceMembersAsync()
    {
        if (!ValidateManage(out var error)) { ManageStatusMessage = error; return; }

        // The inline "yes I am sure" checkbox this used to have sat in the same scroll column as
        // every harmless action and could be ticked minutes before the button was pressed. A modal
        // asks at the moment of the decision and names what is about to be discarded.
        var incoming = ReplaceMembersText
            .Split(new[] { ',', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Length;

        if (!_dialogs.ConfirmDestructive(
                $"Replace the entire membership of {SelectedGroup!.DisplayName}?",
                $"Everyone currently in the group who is not in your list is removed. The group ends up with exactly the {incoming} identities you entered.",
                "Replace membership",
                "The previous membership is written to the log first, so it can be reconstructed."))
            return;

        IsManageBusy = true;
        ManageStatusMessage = "Replacing membership...";
        Transcript.BeginOperation($"Replace membership of {SelectedGroup!.DisplayName}");
        try
        {
            var result = await _groupAdmin.ReplaceMembersAsync(
                BuildListRequest(ListChangeMode.Replace, ReplaceMembersText),
                msg => ManageStatusMessage = msg);

            ReportManage(result, "Membership replaced. The previous list was written to the log.");
            if (result.Succeeded)
            {
                ReplaceMembersText = "";
                await RefreshSelectedGroupAsync();
            }
        }
        catch (Exception ex)
        {
            ManageStatus($"Replace failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    /// <summary>
    /// Rewrites the group's MailTip from its current owners. Repairs tips written before the
    /// flattened-ManagedBy fix, which otherwise only correct themselves on the next owner change.
    /// </summary>
    [RelayCommand]
    private async Task RefreshMailTipAsync()
    {
        if (!ValidateManage(out var error)) { ManageStatusMessage = error; return; }

        IsManageBusy = true;
        ManageStatusMessage = "Rebuilding the MailTip from the current owners...";
        Transcript.BeginOperation($"Refresh MailTip: {SelectedGroup!.DisplayName}");
        try
        {
            var result = await _groupAdmin.RefreshMailTipAsync(
                GroupIdentity(SelectedGroup!), ManageTaskNumber.Trim());

            ManageStatusSeverity = result.Succeeded ? AlertSeverity.Success : AlertSeverity.Error;

            ManageStatusMessage = result.Succeeded
                ? $"MailTip is now: {result.Info}"
                : $"MailTip refresh failed: {result.ErrorMessage}";
        }
        catch (Exception ex)
        {
            ManageStatus($"MailTip refresh failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    [RelayCommand]
    private async Task RenameGroupAsync()
    {
        if (!ValidateManage(out var error)) { ManageStatusMessage = error; return; }

        IsManageBusy = true;
        ManageStatusMessage = "Renaming...";
        Transcript.BeginOperation($"Rename {SelectedGroup!.DisplayName} -> {RenameNewName}");
        try
        {
            var result = await _groupAdmin.RenameAsync(new GroupRenameRequest
            {
                TaskNumber = ManageTaskNumber.Trim(),
                RequesterIdentity = ManageRequesterIdentity.Trim(),
                GroupIdentity = GroupIdentity(SelectedGroup!),
                NewName = RenameNewName.Trim(),
                NewAddress = RenameNewAddress.Trim(),
            }, msg => ManageStatusMessage = msg);

            if (result.Succeeded)
            {
                ManageStatusSeverity = result.Warning is { Length: > 0 } ? AlertSeverity.Warning : AlertSeverity.Success;
                ManageStatusMessage = $"Renamed to '{result.DisplayName}' ({result.Address}). The previous address is kept as an alias."
                    + (result.Warning is { Length: > 0 } w ? $" WARNING: {w}" : "");
                await SearchAsync();
            }
            else
            {
                ManageStatus($"Rename failed: {result.ErrorMessage}", AlertSeverity.Error);
            }
        }
        catch (Exception ex)
        {
            ManageStatus($"Rename failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    [RelayCommand]
    private async Task RemoveGroupAsync()
    {
        if (!ValidateManage(out var error)) { ManageStatusMessage = error; return; }

        if (!_dialogs.ConfirmDestructive(
                $"Delete {SelectedGroup!.DisplayName}?",
                "The group and its membership are deleted from the tenant. This cannot be undone from here - "
                + "restoring it means recreating the group and its members by hand.",
                "Delete group",
                $"{SelectedGroup.PrimarySmtpAddress} - members, owners, aliases and settings are written to the log before anything is deleted."))
            return;

        var name = SelectedGroup!.DisplayName;
        IsManageBusy = true;
        ManageStatusMessage = "Removing...";
        Transcript.BeginOperation($"Remove group: {name}");
        try
        {
            var result = await _groupAdmin.RemoveAsync(new GroupRemovalRequest
            {
                TaskNumber = ManageTaskNumber.Trim(),
                RequesterIdentity = ManageRequesterIdentity.Trim(),
                GroupIdentity = GroupIdentity(SelectedGroup!),
            }, msg => ManageStatusMessage = msg);

            if (result.Succeeded)
            {
                ManageStatusSeverity = AlertSeverity.Success;
                ManageStatusMessage = $"'{name}' removed as {result.RemovedAs}. A snapshot was written to the log."
                    + (result.Warning is { Length: > 0 } w ? $" WARNING: {w}" : "");
                SelectedGroup = null;
                await SearchAsync();
            }
            else
            {
                ManageStatus($"Removal failed: {result.ErrorMessage}", AlertSeverity.Error);
            }
        }
        catch (Exception ex)
        {
            ManageStatus($"Removal failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    /// <summary>Re-reads the selected group's members and settings after a change.</summary>
    private async Task RefreshSelectedGroupAsync()
    {
        var group = SelectedGroup;
        if (group is null)
            return;

        await LoadMembersAsync(group);
        await LoadGroupDetailsAsync(group);
    }
}
