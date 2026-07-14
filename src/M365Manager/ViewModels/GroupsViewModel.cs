using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Core.Exchange;
using M365Manager.Core.M365;
using M365Manager.Data.Logging;

namespace M365Manager.ViewModels;

public sealed partial class GroupsViewModel : ObservableObject
{
    private readonly IExchangeService _exchange;
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

    public GroupsViewModel(IExchangeService exchange, ILogService log, IM365AuthService auth)
    {
        _exchange = exchange;
        _log = log;
        _auth = auth;

        MembersView = CollectionViewSource.GetDefaultView(Members);
        MembersView.Filter = FilterMember;
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
            StatusMessage = "Connected to Exchange Online. Enter a name and search.";
            await SafeLogAsync("Connect", null, Severity.Success, "Connected to Exchange Online.");
        }
        catch (Exception ex)
        {
            DeviceCodeMessage = "";
            DeviceCode = "";
            StatusMessage = $"Connection failed: {ex.Message}";
            await SafeLogAsync("Connect", null, Severity.Error, $"Exchange connect failed: {ex.Message}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    private void ExtractDeviceCode(string prompt)
    {
        // e.g. "...enter the code ABCD1234 to authenticate."
        var match = System.Text.RegularExpressions.Regex.Match(prompt, @"code\s+([A-Za-z0-9][A-Za-z0-9-]{4,})");
        if (match.Success)
            DeviceCode = match.Groups[1].Value;
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

        var match = System.Text.RegularExpressions.Regex.Match(prompt, @"https?://\S+");
        if (!match.Success)
            return;

        _browserOpened = true;
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = match.Value,
                UseShellExecute = true,
            });
        }
        catch
        {
            // If the browser cannot be opened, the URL is still shown in the banner.
        }
    }

    [RelayCommand]
    private async Task SearchAsync()
    {
        if (!IsConnected)
        {
            StatusMessage = "Please connect to Exchange Online first.";
            return;
        }
        if (string.IsNullOrWhiteSpace(SearchText))
        {
            StatusMessage = "Enter a group name, alias or address to search.";
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

            StatusMessage = $"{SearchResults.Count} group(s) found.";
            await SafeLogAsync("Search", SearchText, Severity.Info, $"Group search '{SearchText}' returned {SearchResults.Count}.");
        }
        catch (Exception ex)
        {
            StatusMessage = $"Search failed: {ex.Message}";
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
        _ = LoadMembersAsync(value);
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

            StatusMessage = $"{group.DisplayName}: {Members.Count} member(s).";
            await SafeLogAsync("ViewMembers", group.PrimarySmtpAddress, Severity.Info,
                $"Viewed {Members.Count} members of {group.DisplayName}.");
        }
        catch (Exception ex)
        {
            StatusMessage = $"Could not load members: {ex.Message}";
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
            StatusMessage = "Select a group first.";
            return;
        }

        var identities = NewMemberAddress
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();
        if (identities.Count == 0)
        {
            StatusMessage = "Enter a SamAccountName, email or UPN to add (comma-separate for multiple).";
            return;
        }
        if (!TryRequireTaskNumber(out var task))
            return;

        var group = SelectedGroup;
        IsBusy = true;
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
            StatusMessage = SummarizeOutcome(results, "Added", group.DisplayName, "to");
        }
        catch (Exception ex)
        {
            StatusMessage = $"Add failed: {ex.Message}";
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
            StatusMessage = "Select a group first.";
            return;
        }

        var members = _selectedMembers.Count > 0
            ? _selectedMembers
            : SelectedMember is not null ? new[] { SelectedMember } : Array.Empty<GroupMemberInfo>();
        if (members.Count == 0)
        {
            StatusMessage = "Select one or more members to remove (click a row, or Ctrl/Shift-click for multiple).";
            return;
        }
        if (!TryRequireTaskNumber(out var task))
            return;

        var names = members.Select(m => m.DisplayName).ToList();
        var namesText = names.Count <= 10 ? string.Join(", ", names) : $"{string.Join(", ", names.Take(10))}, and {names.Count - 10} more";
        var confirm = System.Windows.MessageBox.Show(
            $"Remove {members.Count} member(s) from {SelectedGroup.DisplayName}?\n\n{namesText}",
            "Confirm removal",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning);

        if (confirm != System.Windows.MessageBoxResult.Yes)
            return;

        var group = SelectedGroup;
        var identities = members.Select(m => string.IsNullOrWhiteSpace(m.PrimarySmtpAddress) ? m.DisplayName : m.PrimarySmtpAddress).ToList();
        IsBusy = true;
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
            StatusMessage = SummarizeOutcome(results, "Removed", group.DisplayName, "from");
        }
        catch (Exception ex)
        {
            StatusMessage = $"Remove failed: {ex.Message}";
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

        StatusMessage = "Enter the ServiceNow task number authorizing this change before continuing.";
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
            StatusMessage = "Select a group first.";
            return;
        }

        var group = SelectedGroup;
        IsBusy = true;
        StatusMessage = $"Syncing members of {group.DisplayName} from Exchange...";
        try
        {
            var members = await _exchange.SyncMembersAsync(group);
            Members.Clear();
            foreach (var m in members)
                Members.Add(m);

            StatusMessage = $"{group.DisplayName}: {Members.Count} direct member(s) synced.";
            await SafeLogAsync("SyncMembers", group.PrimarySmtpAddress, Severity.Success,
                $"Synced {Members.Count} direct members of {group.DisplayName} from Exchange.");
        }
        catch (Exception ex)
        {
            StatusMessage = $"Sync failed: {ex.Message}";
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
            StatusMessage = "Select a group first.";
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
            StatusMessage = $"Exported to {dialog.FileName}.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Export failed: {ex.Message}";
        }
    }

    [RelayCommand]
    private void ExportToCsv()
    {
        if (SelectedGroup is null)
        {
            StatusMessage = "Select a group first.";
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
            StatusMessage = $"Exported to {dialog.FileName}.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Export failed: {ex.Message}";
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
}
