using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Controls;
using M365Manager.Core.PowerShell;
using M365Manager.Services;
using M365Manager.Core.RoomResources;

namespace M365Manager.ViewModels;

public sealed partial class RoomResourcesViewModel : ObservableObject, ITabbedPage
{
    /// <summary>Which tab the view shows; set by the shell when a Home tile navigates here.</summary>
    [ObservableProperty] private int _selectedTabIndex;

    private readonly IRoomResourceService _rooms;
    private readonly IDialogService _dialogs;

    // --- Create form ---
    [ObservableProperty] private string _taskNumber = "";
    [ObservableProperty] private string _rawName = "";
    [ObservableProperty] private ResourceKind _kind = ResourceKind.Room;
    [ObservableProperty] private RoomAccessModel _accessModel = RoomAccessModel.GeneralUseSiteDelegates;
    [ObservableProperty] private BookingPolicyPreset _preset = BookingPolicyPreset.Standard;
    [ObservableProperty] private string _capacity = "";
    [ObservableProperty] private string _building = "";
    [ObservableProperty] private string _floor = "";
    [ObservableProperty] private string _delegateMembers = "";
    [ObservableProperty] private string _userMembers = "";

    /// <summary>Who gets the confirmation e-mail. Empty means none is sent - see RoomResourceCreationRequest.</summary>
    [ObservableProperty] private string _requesterIdentity = "";
    [ObservableProperty] private string _mailboxDomainOverride = "";
    [ObservableProperty] private string _groupDomainOverride = "";

    // Custom booking policy values, only used when Preset == Custom.
    [ObservableProperty] private string _customBookingWindowDays = "90";
    [ObservableProperty] private string _customMaxDurationMinutes = "1440";
    [ObservableProperty] private bool _customAllowRecurring = true;

    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = "Fill in the fields and click Create.";

    // --- Live preview ---
    [ObservableProperty] private string _previewDisplayName = "";
    [ObservableProperty] private string _previewAddress = "";
    [ObservableProperty] private string _previewRoomList = "";
    [ObservableProperty] private string _previewDelegateGroup = "";
    [ObservableProperty] private string _previewUsersGroup = "";
    [ObservableProperty] private string _previewSiteInfo = "";
    [ObservableProperty] private string _previewProblems = "";

    /// <summary>Guards against an older preview lookup overwriting a newer one.</summary>
    private int _previewRequestVersion;

    public bool IsRoom => Kind == ResourceKind.Room;
    public bool IsRestricted => AccessModel == RoomAccessModel.Restricted;
    public bool IsCustomPolicy => Preset == BookingPolicyPreset.Custom;

    public IReadOnlyList<ResourceKind> ResourceKinds { get; } = Enum.GetValues<ResourceKind>();
    public IReadOnlyList<RoomAccessModel> AccessModels { get; } = Enum.GetValues<RoomAccessModel>();
    public IReadOnlyList<BookingPolicyPreset> Presets { get; } = Enum.GetValues<BookingPolicyPreset>();
    public IReadOnlyList<MembershipChangeMode> MembershipModes { get; } = Enum.GetValues<MembershipChangeMode>();
    public IReadOnlyList<RoomGroupKind> GroupKinds { get; } = Enum.GetValues<RoomGroupKind>();

    // --- Overview grid ---
    [ObservableProperty] private string _overviewFilterText = "";
    [ObservableProperty] private bool _isOverviewBusy;
    [ObservableProperty] private string _overviewStatusMessage = "";
    [ObservableProperty] private RoomResourceOverviewRow? _selectedRoom;

    public ObservableCollection<RoomResourceOverviewRow> Rooms { get; } = new();
    public ICollectionView RoomsView { get; }

    // --- Manage panel (edit details / membership / remove) ---
    [ObservableProperty] private string _manageTaskNumber = "";

    /// <summary>Requester for the manage panel's operations (edit / membership / remove).</summary>
    [ObservableProperty] private string _manageRequesterIdentity = "";
    [ObservableProperty] private string _editCapacity = "";
    [ObservableProperty] private string _editTimeZone = "";
    [ObservableProperty] private bool _changeBookingPolicy;
    [ObservableProperty] private BookingPolicyPreset _editPreset = BookingPolicyPreset.Standard;
    [ObservableProperty] private bool _isManageBusy;
    [ObservableProperty] private string _manageStatusMessage = "";

    [ObservableProperty] private RoomGroupKind _membershipGroupKind = RoomGroupKind.Delegates;
    [ObservableProperty] private MembershipChangeMode _membershipMode = MembershipChangeMode.Add;
    [ObservableProperty] private string _membershipIdentities = "";
    [ObservableProperty] private string _currentMembers = "";

    [ObservableProperty] private bool _removeOrphanedGroups = true;

    // Status severity, so a failure stops looking like a success.
    [ObservableProperty] private AlertSeverity _statusSeverity = AlertSeverity.Info;
    [ObservableProperty] private AlertSeverity _overviewStatusSeverity = AlertSeverity.Info;
    [ObservableProperty] private AlertSeverity _manageStatusSeverity = AlertSeverity.Info;

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

    private void ManageStatus(string text, AlertSeverity severity = AlertSeverity.Info)
    {
        ManageStatusSeverity = severity;
        ManageStatusMessage = text;
    }

    /// <summary>Backs the PowerShell console on this page - see PowerShellConsole.</summary>
    public IPowerShellTranscript Transcript { get; }

    public RoomResourcesViewModel(IRoomResourceService rooms, IPowerShellTranscript transcript, IDialogService dialogs)
    {
        _rooms = rooms;
        _dialogs = dialogs;
        Transcript = transcript;

        RoomsView = CollectionViewSource.GetDefaultView(Rooms);
        RoomsView.Filter = FilterRow;

        _ = RefreshOverviewAsync();
    }

    // --- preview plumbing ---

    partial void OnRawNameChanged(string value) => _ = RefreshPreviewAsync();
    partial void OnBuildingChanged(string value) => _ = RefreshPreviewAsync();
    partial void OnFloorChanged(string value) => _ = RefreshPreviewAsync();
    partial void OnCapacityChanged(string value) => _ = RefreshPreviewAsync();
    partial void OnMailboxDomainOverrideChanged(string value) => _ = RefreshPreviewAsync();
    partial void OnGroupDomainOverrideChanged(string value) => _ = RefreshPreviewAsync();

    // A selection is one deliberate change, not a burst of keystrokes - no need to wait.
    partial void OnKindChanged(ResourceKind value)
    {
        OnPropertyChanged(nameof(IsRoom));
        _ = RefreshPreviewAsync(debounce: false);
    }

    partial void OnAccessModelChanged(RoomAccessModel value)
    {
        OnPropertyChanged(nameof(IsRestricted));
        _ = RefreshPreviewAsync(debounce: false);
    }

    partial void OnPresetChanged(BookingPolicyPreset value) => OnPropertyChanged(nameof(IsCustomPolicy));

    private RoomResourceCreationRequest BuildRequest() => new()
    {
        TaskNumber = TaskNumber.Trim(),
        RequesterIdentity = RequesterIdentity.Trim(),
        RawName = RawName.Trim(),
        Kind = Kind,
        AccessModel = AccessModel,
        Preset = Preset,
        CustomPolicy = Preset == BookingPolicyPreset.Custom ? BuildCustomPolicy() : null,
        Capacity = Capacity.Trim(),
        Building = Building.Trim(),
        Floor = Floor.Trim(),
        DelegateMembers = DelegateMembers.Trim(),
        UserMembers = UserMembers.Trim(),
        MailboxDomain = MailboxDomainOverride.Trim(),
        GroupDomain = GroupDomainOverride.Trim(),
    };

    private BookingPolicy BuildCustomPolicy() => new(
        BookingWindowInDays: int.TryParse(CustomBookingWindowDays.Trim(), out var w) && w > 0 ? w : 90,
        MaximumDurationInMinutes: int.TryParse(CustomMaxDurationMinutes.Trim(), out var d) && d > 0 ? d : 1440,
        AllowRecurringMeetings: CustomAllowRecurring,
        MaximumConflictInstances: CustomAllowRecurring ? 3 : 0,
        ConflictPercentageAllowed: CustomAllowRecurring ? 20 : 0);

    /// <summary>Pause after the last keystroke before the preview asks Exchange anything.</summary>
    private static readonly TimeSpan PreviewDebounce = TimeSpan.FromMilliseconds(500);

    private async Task RefreshPreviewAsync(bool debounce = true)
    {
        var version = ++_previewRequestVersion;

        if (string.IsNullOrWhiteSpace(RawName))
        {
            ClearPreview();
            return;
        }

        try
        {
            // The preview probes Exchange for the mailbox, room list and groups - several
            // cmdlets per call. Without the pause, typing "FRK Cristian Test" ran them for every
            // letter and filled the PowerShell console. A newer keystroke bumps the version and
            // this call gives up before it has touched Exchange.
            if (debounce)
            {
                await Task.Delay(PreviewDebounce);
                if (version != _previewRequestVersion)
                    return;
            }

            var preview = await _rooms.PreviewAsync(BuildRequest());
            if (version != _previewRequestVersion)
                return;

            PreviewDisplayName = preview.Names.DisplayName;
            PreviewAddress = preview.Names.Address;
            PreviewRoomList = Kind == ResourceKind.Room
                ? Decorate(preview.Names.RoomListName, preview.RoomListExists)
                : "";
            PreviewDelegateGroup = Decorate(preview.Names.DelegateGroup, preview.DelegateGroupExists);
            PreviewUsersGroup = preview.Names.UsersGroup.Length > 0
                ? Decorate(preview.Names.UsersGroup, preview.UsersGroupExists)
                : "";
            PreviewSiteInfo = preview.TimeZone is { Length: > 0 }
                ? $"{preview.Names.SiteCode}: {preview.TimeZone}" +
                  (string.IsNullOrWhiteSpace(preview.RegionalAdminGroup) ? "" : $" · {preview.RegionalAdminGroup}")
                : "";
            PreviewProblems = string.Join("\n", preview.Errors);
        }
        catch
        {
            // Preview is best-effort (SQL/Exchange may not be reachable yet) - keep the last value.
        }
    }

    /// <summary>Marks each derived object as reused, newly created, or not yet checkable.</summary>
    private static string Decorate(string name, bool? exists)
    {
        if (name.Length == 0)
            return "";

        var state = exists switch
        {
            true => "exists - will be reused",
            false => "will be created",
            null => "could not check - not connected?",
        };
        return $"{name}  ({state})";
    }

    private void ClearPreview()
    {
        PreviewDisplayName = "";
        PreviewAddress = "";
        PreviewRoomList = "";
        PreviewDelegateGroup = "";
        PreviewUsersGroup = "";
        PreviewSiteInfo = "";
        PreviewProblems = "";
    }

    // --- create ---

    [RelayCommand]
    private async Task CreateAsync()
    {
        if (string.IsNullOrWhiteSpace(TaskNumber))
        {
            Status("Enter the ticket/task number authorizing this creation.", AlertSeverity.Error);
            return;
        }
        if (string.IsNullOrWhiteSpace(RawName))
        {
            StatusMessage = "Enter the name as \"<SITE> <Name>\", e.g. \"NBK Meeting Room 1\".";
            return;
        }

        IsBusy = true;
        StatusMessage = "Starting creation...";
        Transcript.BeginOperation($"Create {Kind}: {PreviewDisplayName}");
        try
        {
            var result = await _rooms.CreateAsync(BuildRequest(), msg => StatusMessage = msg);
            StatusSeverity = !result.Succeeded ? AlertSeverity.Error
                : result.WarningMessage is null ? AlertSeverity.Success
                : AlertSeverity.Warning;
            StatusMessage = result.Succeeded
                ? $"'{result.DisplayName}' ({result.PrimarySmtpAddress}) created successfully."
                  + (result.WarningMessage is null ? "" : $" WARNING: {result.WarningMessage}")
                : $"Creation failed: {result.ErrorMessage}";

            if (result.Succeeded)
            {
                ClearForm();
                _ = RefreshOverviewAsync();
            }
        }
        catch (Exception ex)
        {
            Status($"Creation failed: {ex.Message}", AlertSeverity.Error);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private void ClearForm()
    {
        TaskNumber = "";
        RequesterIdentity = "";
        RawName = "";
        Capacity = "";
        Building = "";
        Floor = "";
        DelegateMembers = "";
        UserMembers = "";
        ClearPreview();
    }

    // --- overview ---

    partial void OnOverviewFilterTextChanged(string value) => RoomsView.Refresh();

    private bool FilterRow(object obj)
    {
        if (string.IsNullOrWhiteSpace(OverviewFilterText))
            return true;
        if (obj is not RoomResourceOverviewRow r)
            return true;

        var term = OverviewFilterText.Trim();
        return Contains(r.DisplayName, term)
            || Contains(r.Alias, term)
            || Contains(r.PrimarySmtpAddress, term)
            || Contains(r.SiteCode, term)
            || Contains(r.Office, term)
            || Contains(r.RoomListName, term)
            || Contains(r.DelegateGroup, term);
    }

    private static bool Contains(string? value, string term)
        => !string.IsNullOrEmpty(value) && value.Contains(term, StringComparison.OrdinalIgnoreCase);

    [RelayCommand]
    private async Task RefreshOverviewAsync()
    {
        IsOverviewBusy = true;
        OverviewStatusMessage = "Loading rooms and resources from SQL...";
        try
        {
            var rows = await _rooms.GetOverviewAsync();
            Rooms.Clear();
            foreach (var row in rows)
                Rooms.Add(row);

            OverviewStatus($"{Rooms.Count} room(s)/resource(s) loaded.", Rooms.Count > 0 ? AlertSeverity.Success : AlertSeverity.Info);
        }
        catch (Exception ex)
        {
            OverviewStatus($"Could not load rooms: {ex.Message}", AlertSeverity.Error);
        }
        finally
        {
            IsOverviewBusy = false;
        }
    }

    [RelayCommand]
    private void ExportToExcel()
    {
        var filtered = RoomsView.Cast<RoomResourceOverviewRow>().ToList();
        if (filtered.Count == 0)
        {
            OverviewStatus("Nothing to export - the (filtered) list is empty.", AlertSeverity.Error);
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel Workbook (*.xlsx)|*.xlsx",
            FileName = "RoomsAndResources.xlsx",
        };
        if (dialog.ShowDialog() != true)
            return;

        try
        {
            RoomResourceExportService.ExportToExcel(dialog.FileName, filtered);
            OverviewStatus($"Exported {filtered.Count} row(s) to {dialog.FileName}.", AlertSeverity.Success);
        }
        catch (Exception ex)
        {
            OverviewStatus($"Export failed: {ex.Message}", AlertSeverity.Error);
        }
    }

    /// <summary>Selecting a row pre-fills the manage panel, so nothing has to be retyped.</summary>
    partial void OnSelectedRoomChanged(RoomResourceOverviewRow? value)
    {
        CurrentMembers = "";
        ManageStatusMessage = "";

        if (value is null)
            return;

        EditCapacity = value.Capacity?.ToString() ?? "";
        EditTimeZone = value.TimeZone;
        ChangeBookingPolicy = false;
        if (Enum.TryParse<BookingPolicyPreset>(value.BookingPolicy, out var preset))
            EditPreset = preset;

        // A general-use room has no users group - default the membership panel to what it has.
        MembershipGroupKind = value.HasUsersGroup ? MembershipGroupKind : RoomGroupKind.Delegates;
        _ = LoadCurrentMembersAsync();

        // The rename form starts from the room as it is, so a ticket that changes one thing only
        // needs that one thing typed.
        _applyingRenamePrefill = true;
        try
        {
            RenameRawName = value.DisplayName;
            RenameBuilding = value.Building;
            RenameFloor = value.Floor;
            RenameCapacity = value.Capacity?.ToString() ?? "";
        }
        finally
        {
            _applyingRenamePrefill = false;
        }
        InvalidateRenamePlan();
    }

    partial void OnMembershipGroupKindChanged(RoomGroupKind value) => _ = LoadCurrentMembersAsync();

    private string? SelectedGroupIdentity => SelectedRoom is null
        ? null
        : MembershipGroupKind == RoomGroupKind.Delegates ? SelectedRoom.DelegateGroup : SelectedRoom.UsersGroup;

    private async Task LoadCurrentMembersAsync()
    {
        var group = SelectedGroupIdentity;
        if (string.IsNullOrWhiteSpace(group))
        {
            CurrentMembers = MembershipGroupKind == RoomGroupKind.Users
                ? "(general-use room - no authorized-users group)"
                : "";
            return;
        }

        try
        {
            var members = await _rooms.GetGroupMembersAsync(group);
            CurrentMembers = members.Count > 0 ? string.Join("\n", members) : "(no members)";
        }
        catch (Exception ex)
        {
            CurrentMembers = $"(could not read members: {ex.Message})";
        }
    }

    // --- manage: edit details ---

    [RelayCommand]
    private async Task UpdateDetailsAsync()
    {
        if (SelectedRoom is null)
        {
            ManageStatus("Select a room or resource in the grid first.", AlertSeverity.Error);
            return;
        }
        if (string.IsNullOrWhiteSpace(ManageTaskNumber))
        {
            ManageStatus("Enter the ticket/task number authorizing this change.", AlertSeverity.Error);
            return;
        }

        var request = new RoomDetailsUpdateRequest
        {
            TaskNumber = ManageTaskNumber.Trim(),
            RequesterIdentity = ManageRequesterIdentity.Trim(),
            Address = SelectedRoom.PrimarySmtpAddress,
            Kind = string.Equals(SelectedRoom.ResourceKind, "Equipment", StringComparison.OrdinalIgnoreCase)
                ? ResourceKind.Equipment
                : ResourceKind.Room,
            Capacity = EditCapacity.Trim(),
            TimeZone = EditTimeZone.Trim(),
            Preset = ChangeBookingPolicy ? EditPreset : null,
            CustomPolicy = ChangeBookingPolicy && EditPreset == BookingPolicyPreset.Custom ? BuildCustomPolicy() : null,
        };

        IsManageBusy = true;
        ManageStatusMessage = "Applying changes...";
        Transcript.BeginOperation($"Update details: {SelectedRoom.DisplayName}");
        try
        {
            var result = await _rooms.UpdateDetailsAsync(request, msg => ManageStatusMessage = msg);
            ManageStatusSeverity = !result.Succeeded ? AlertSeverity.Error
                : result.WarningMessage is null ? AlertSeverity.Success
                : AlertSeverity.Warning;
            ManageStatusMessage = result.Succeeded
                ? $"Updated {result.PrimarySmtpAddress}."
                  + (result.WarningMessage is null ? "" : $" WARNING: {result.WarningMessage}")
                : $"Update failed: {result.ErrorMessage}";

            if (result.Succeeded)
                _ = RefreshOverviewAsync();
        }
        catch (Exception ex)
        {
            ManageStatus($"Update failed: {ex.Message}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    // --- manage: membership ---

    [RelayCommand]
    private async Task ChangeMembershipAsync()
    {
        if (SelectedRoom is null)
        {
            ManageStatus("Select a room or resource in the grid first.", AlertSeverity.Error);
            return;
        }
        if (string.IsNullOrWhiteSpace(ManageTaskNumber))
        {
            ManageStatus("Enter the ticket/task number authorizing this change.", AlertSeverity.Error);
            return;
        }
        if (string.IsNullOrWhiteSpace(MembershipIdentities))
        {
            ManageStatus("Enter at least one identity to add or remove.", AlertSeverity.Error);
            return;
        }

        var request = new RoomMembershipChangeRequest
        {
            TaskNumber = ManageTaskNumber.Trim(),
            RequesterIdentity = ManageRequesterIdentity.Trim(),
            Address = SelectedRoom.PrimarySmtpAddress,
            GroupIdentity = SelectedGroupIdentity ?? "",
            GroupKind = MembershipGroupKind,
            Mode = MembershipMode,
            Identities = MembershipIdentities.Trim(),
        };

        IsManageBusy = true;
        ManageStatusMessage = "Applying membership change...";
        Transcript.BeginOperation($"{MembershipMode} {MembershipGroupKind} member(s): {SelectedRoom.DisplayName}");
        try
        {
            var result = await _rooms.ChangeMembershipAsync(request);
            if (!result.Succeeded && result.Results.Count == 0)
            {
                ManageStatus($"Membership change failed: {result.ErrorMessage}", AlertSeverity.Error);
            }
            else
            {
                var ok = result.Results.Count(r => r.Succeeded);
                var failed = result.Results.Where(r => !r.Succeeded).ToList();
                ManageStatusMessage = (failed.Count == 0
                    ? $"{ok} identity/identities {(MembershipMode == MembershipChangeMode.Add ? "added" : "removed")}."
                    : $"{ok} succeeded, {failed.Count} failed: " +
                      string.Join("; ", failed.Select(f => $"{f.Identity} ({f.Error})")))
                    + (result.WarningMessage is null ? "" : $" WARNING: {result.WarningMessage}");

                MembershipIdentities = "";
                await LoadCurrentMembersAsync();
            }
        }
        catch (Exception ex)
        {
            ManageStatus($"Membership change failed: {ex.Message}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    // --- manage: remove ---

    [RelayCommand]
    private async Task RemoveAsync()
    {
        if (SelectedRoom is null)
        {
            ManageStatus("Select a room or resource in the grid first.", AlertSeverity.Error);
            return;
        }
        if (string.IsNullOrWhiteSpace(ManageTaskNumber))
        {
            ManageStatus("Enter the ticket/task number authorizing this removal.", AlertSeverity.Error);
            return;
        }
        // Replaces the inline "yes I am sure" checkbox: a modal asks at the moment of the
        // decision and spells out whether the room's own groups are going with it.
        var extras = RemoveOrphanedGroups
            ? "Its own delegate/users groups are deleted too, and it is dropped from its room list. "
              + "Shared site-wide delegate groups are never deleted."
            : "Its delegate/users groups and its room list membership are left in place.";

        if (!_dialogs.ConfirmDestructive(
                $"Delete {SelectedRoom.DisplayName}?",
                "The mailbox is deleted from the tenant. This cannot be undone from here.",
                "Delete mailbox",
                $"{SelectedRoom.PrimarySmtpAddress}\n\n{extras}\n\n"
                + "A full snapshot - mailbox, statistics, permissions, calendar processing and configuration - "
                + "is written to the log before anything is deleted."))
            return;

        var address = SelectedRoom.PrimarySmtpAddress;
        var request = new RoomRemovalRequest
        {
            TaskNumber = ManageTaskNumber.Trim(),
            RequesterIdentity = ManageRequesterIdentity.Trim(),
            Address = address,
            RemoveOrphanedGroups = RemoveOrphanedGroups,
        };

        IsManageBusy = true;
        ManageStatusMessage = "Removing...";
        Transcript.BeginOperation($"Remove: {address}");
        try
        {
            var result = await _rooms.RemoveAsync(request, msg => ManageStatusMessage = msg);
            if (result.Succeeded)
            {
                var groups = result.RemovedGroups.Count > 0
                    ? $" Removed groups: {string.Join(", ", result.RemovedGroups)}."
                    : "";
                ManageStatusMessage = $"{address} removed.{groups} A full pre-removal snapshot was written to the log."
                    + (result.WarningMessage is null ? "" : $" WARNING: {result.WarningMessage}");
                _ = RefreshOverviewAsync();
            }
            else
            {
                ManageStatus($"Removal failed: {result.ErrorMessage}", AlertSeverity.Error);
            }
        }
        catch (Exception ex)
        {
            ManageStatus($"Removal failed: {ex.Message}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    // --- manage: rename ---
    //
    // Two steps, like the shared mailbox rename: "Preview rename" reads the room from Exchange and
    // works out everything that would change - address, groups, room list, time zone, admins -
    // and only then does the Rename button do anything. A rename touches up to a dozen objects.

    [ObservableProperty] private string _renameRawName = "";
    [ObservableProperty] private string _renameBuilding = "";
    [ObservableProperty] private string _renameFloor = "";
    [ObservableProperty] private string _renameCapacity = "";
    [ObservableProperty] private string _renamePlanSummary = "";
    [ObservableProperty] private bool _hasRenamePlan;
    private RoomRenamePlan? _renamePlan;

    /// <summary>Set while a room selection fills the fields, so that doesn't count as an edit.</summary>
    private bool _applyingRenamePrefill;

    // Any change to the inputs makes the previewed plan stale.
    partial void OnRenameRawNameChanged(string value) => OnRenameInputChanged();
    partial void OnRenameBuildingChanged(string value) => OnRenameInputChanged();
    partial void OnRenameFloorChanged(string value) => OnRenameInputChanged();
    partial void OnRenameCapacityChanged(string value) => OnRenameInputChanged();

    private void OnRenameInputChanged()
    {
        if (!_applyingRenamePrefill)
            InvalidateRenamePlan();
    }

    private void InvalidateRenamePlan()
    {
        _renamePlan = null;
        HasRenamePlan = false;
        RenamePlanSummary = "";
    }

    private RoomRenameRequest BuildRenameRequest() => new()
    {
        TaskNumber = ManageTaskNumber.Trim(),
        RequesterIdentity = ManageRequesterIdentity.Trim(),
        MailboxIdentity = SelectedRoom?.PrimarySmtpAddress ?? "",
        RawName = RenameRawName.Trim(),
        Building = RenameBuilding.Trim(),
        Floor = RenameFloor.Trim(),
        Capacity = RenameCapacity.Trim(),
    };

    [RelayCommand]
    private async Task PreviewRenameAsync()
    {
        if (SelectedRoom is null)
        {
            ManageStatus("Select a room or resource in the grid first.", AlertSeverity.Error);
            return;
        }

        IsManageBusy = true;
        InvalidateRenamePlan();
        ManageStatus("Reading the room and working out the rename...");
        Transcript.BeginOperation($"Preview rename: {SelectedRoom.PrimarySmtpAddress}");
        try
        {
            var current = await _rooms.LookUpForRenameAsync(SelectedRoom.PrimarySmtpAddress);
            var plan = await _rooms.PlanRenameAsync(current, BuildRenameRequest());
            RenamePlanSummary = DescribeRenamePlan(plan);

            if (plan.Errors.Count > 0)
            {
                ManageStatus(string.Join(" ", plan.Errors), AlertSeverity.Error);
                return;
            }

            _renamePlan = plan;
            HasRenamePlan = true;
            ManageStatus("Check the changes below, then rename.");
        }
        catch (Exception ex)
        {
            ManageStatus($"Preview failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    private static string DescribeRenamePlan(RoomRenamePlan plan)
    {
        var lines = new List<string>
        {
            $"Name: {plan.Current.DisplayName}  ->  {plan.NewNames.DisplayName}",
            $"Address: {plan.Current.PrimarySmtpAddress}  ->  {plan.NewNames.Address}",
        };
        foreach (var (from, to) in plan.GroupRenames)
            lines.Add($"Group: {from}  ->  {to}");
        if (plan.DelegateSwap is { } swap)
            lines.Add($"Delegates: {(swap.From.Length > 0 ? swap.From : "(none)")}  ->  {swap.To}");
        if (plan.RoomListMove is { } move)
            lines.Add($"Room list: {move.From}  ->  {move.To}");
        if (plan.NewTimeZone is { } tz)
            lines.Add($"Time zone: {(plan.Current.TimeZone.Length > 0 ? plan.Current.TimeZone : "(not set)")}  ->  {tz}");
        if (plan.NewRegionalAdminGroup is { } admins)
            lines.Add($"Regional admins: {admins}");
        foreach (var note in plan.Notes)
            lines.Add(note);
        return string.Join("\n", lines);
    }

    [RelayCommand]
    private async Task RenameRoomAsync()
    {
        if (SelectedRoom is null)
        {
            ManageStatus("Select a room or resource in the grid first.", AlertSeverity.Error);
            return;
        }
        if (string.IsNullOrWhiteSpace(ManageTaskNumber))
        {
            ManageStatus("Enter the ticket/task number authorizing this rename.", AlertSeverity.Error);
            return;
        }
        if (_renamePlan is not { } plan)
        {
            ManageStatus("Preview the rename first - it shows everything that changes with it.", AlertSeverity.Error);
            return;
        }

        if (!_dialogs.ConfirmDestructive(
                $"Rename {plan.Current.DisplayName}?",
                $"It becomes '{plan.NewNames.DisplayName}' ({plan.NewNames.Address}).",
                "Rename room",
                RenamePlanSummary))
            return;

        IsManageBusy = true;
        ManageStatus("Starting rename...");
        Transcript.BeginOperation($"Rename: {plan.Current.DisplayName} -> {plan.NewNames.DisplayName}");
        try
        {
            var result = await _rooms.RenameAsync(BuildRenameRequest(), msg => ManageStatusMessage = msg);
            if (result.Succeeded)
            {
                var changes = result.Changes.Count == 0 ? "" : $" {string.Join("; ", result.Changes)}.";
                ManageStatus(
                    $"'{result.PreviousDisplayName}' renamed to '{result.DisplayName}' ({result.PrimarySmtpAddress}).{changes}"
                    + (result.WarningMessage is null ? "" : $" WARNING: {result.WarningMessage}"),
                    result.WarningMessage is null ? AlertSeverity.Success : AlertSeverity.Warning);
                InvalidateRenamePlan();
                _ = RefreshOverviewAsync();
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
}
