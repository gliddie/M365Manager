using System.Collections.ObjectModel;
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

    private bool _browserOpened;

    public ObservableCollection<DistributionGroupInfo> SearchResults { get; } = new();
    public ObservableCollection<GroupMemberInfo> Members { get; } = new();

    public GroupsViewModel(IExchangeService exchange, ILogService log, IM365AuthService auth)
    {
        _exchange = exchange;
        _log = log;
        _auth = auth;
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

        IsBusy = true;
        StatusMessage = $"Searching for '{SearchText}'...";
        try
        {
            var groups = await _exchange.SearchGroupsAsync(SearchText);
            SearchResults.Clear();
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

    partial void OnSelectedGroupChanged(DistributionGroupInfo? value)
    {
        OwnersText = value is null ? "" : string.Join(", ", value.ManagedBy);
        _ = LoadMembersAsync(value);
    }

    private async Task LoadMembersAsync(DistributionGroupInfo? group)
    {
        Members.Clear();
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
        if (string.IsNullOrWhiteSpace(NewMemberAddress))
        {
            StatusMessage = "Enter a member email or UPN to add.";
            return;
        }

        var member = NewMemberAddress.Trim();
        IsBusy = true;
        try
        {
            await _exchange.AddMemberAsync(SelectedGroup, member);
            await SafeLogAsync("AddMember", SelectedGroup.PrimarySmtpAddress, Severity.Success,
                $"Added {member} to {SelectedGroup.DisplayName}.");
            NewMemberAddress = "";
            await LoadMembersAsync(SelectedGroup);
            StatusMessage = $"Added {member} to {SelectedGroup.DisplayName}.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Add failed: {ex.Message}";
            await SafeLogAsync("AddMember", SelectedGroup.PrimarySmtpAddress, Severity.Error,
                $"Add {member} failed: {ex.Message}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private async Task RemoveMemberAsync()
    {
        if (SelectedGroup is null || SelectedMember is null)
        {
            StatusMessage = "Select a member to remove.";
            return;
        }

        var member = SelectedMember;
        var confirm = System.Windows.MessageBox.Show(
            $"Remove {member.DisplayName} ({member.PrimarySmtpAddress}) from {SelectedGroup.DisplayName}?",
            "Confirm removal",
            System.Windows.MessageBoxButton.YesNo,
            System.Windows.MessageBoxImage.Warning);

        if (confirm != System.Windows.MessageBoxResult.Yes)
            return;

        var identity = string.IsNullOrWhiteSpace(member.PrimarySmtpAddress) ? member.DisplayName : member.PrimarySmtpAddress;
        IsBusy = true;
        try
        {
            await _exchange.RemoveMemberAsync(SelectedGroup, identity);
            await SafeLogAsync("RemoveMember", SelectedGroup.PrimarySmtpAddress, Severity.Success,
                $"Removed {identity} from {SelectedGroup.DisplayName}.");
            await LoadMembersAsync(SelectedGroup);
            StatusMessage = $"Removed {member.DisplayName} from {SelectedGroup.DisplayName}.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Remove failed: {ex.Message}";
            await SafeLogAsync("RemoveMember", SelectedGroup.PrimarySmtpAddress, Severity.Error,
                $"Remove {identity} failed: {ex.Message}");
        }
        finally
        {
            IsBusy = false;
        }
    }

    private async Task SafeLogAsync(string action, string? target, Severity severity, string message)
    {
        try
        {
            await _log.WriteAsync(new LogEntry
            {
                Area = "Groups",
                Action = action,
                TargetObject = target,
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
