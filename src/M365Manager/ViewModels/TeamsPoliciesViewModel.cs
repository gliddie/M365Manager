using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Core.PowerShell;
using M365Manager.Core.TeamsPolicies;

namespace M365Manager.ViewModels;

/// <summary>
/// One configured prefix as the page shows it: what the user holds, what they could hold, and the
/// pending choice. Each category is independent - changing one never affects another.
/// </summary>
public sealed partial class PolicyCategoryViewModel : ObservableObject
{
    public required string Prefix { get; init; }

    public ObservableCollection<TeamsPolicyGroup> AvailableGroups { get; } = new();

    /// <summary>What the user holds right now, as text. "(none)" when they hold nothing.</summary>
    [ObservableProperty] private string _currentText = "(none)";

    /// <summary>More than one group under this prefix - a state the one-per-prefix rule should prevent.</summary>
    [ObservableProperty] private bool _hasConflict;

    [ObservableProperty] private TeamsPolicyGroup? _selectedGroup;

    /// <summary>Set once Apply is clicked with a change pending; cleared when confirmed or cancelled.</summary>
    [ObservableProperty] private string _pendingConfirmation = "";

    public IReadOnlyList<TeamsPolicyGroup> CurrentGroups { get; set; } = Array.Empty<TeamsPolicyGroup>();

    public bool HasPendingConfirmation => PendingConfirmation.Length > 0;

    partial void OnPendingConfirmationChanged(string value) => OnPropertyChanged(nameof(HasPendingConfirmation));
}

public sealed partial class TeamsPoliciesViewModel : ObservableObject
{
    private readonly ITeamsPolicyService _policies;

    /// <summary>
    /// Backs the console on this page. This feature is pure Graph - without the console showing
    /// Graph requests it would look as though nothing happened at all.
    /// </summary>
    public IPowerShellTranscript Transcript { get; }

    [ObservableProperty] private string _userIdentity = "";
    [ObservableProperty] private string _taskNumber = "";
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string _statusMessage = "Enter a user and click Look up.";
    [ObservableProperty] private string _userSummary = "";

    private PolicyUser? _user;

    public ObservableCollection<PolicyCategoryViewModel> Categories { get; } = new();

    public TeamsPoliciesViewModel(ITeamsPolicyService policies, IPowerShellTranscript transcript)
    {
        _policies = policies;
        Transcript = transcript;
    }

    [RelayCommand]
    private async Task LookUpAsync()
    {
        if (string.IsNullOrWhiteSpace(UserIdentity))
        {
            StatusMessage = "Enter a SamAccountName, UPN or e-mail address.";
            return;
        }

        IsBusy = true;
        StatusMessage = "Looking up...";
        Transcript.BeginOperation($"Look up policy groups: {UserIdentity.Trim()}");
        Categories.Clear();
        UserSummary = "";
        _user = null;

        try
        {
            var lookup = await _policies.LookupUserAsync(UserIdentity.Trim());
            if (!lookup.Succeeded || lookup.User is null)
            {
                StatusMessage = lookup.ErrorMessage ?? "Lookup failed.";
                return;
            }

            _user = lookup.User;
            UserSummary = $"{lookup.User.DisplayName} ({lookup.User.Upn})";

            foreach (var category in lookup.Categories)
                Categories.Add(BuildCategory(category));

            StatusMessage = Categories.Count == 0
                ? "No policy group prefixes are configured - set them under Settings > Teams policies."
                : $"{Categories.Count} categor{(Categories.Count == 1 ? "y" : "ies")} loaded.";
        }
        catch (Exception ex)
        {
            StatusMessage = $"Lookup failed: {ErrorText.Describe(ex)}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    private static PolicyCategoryViewModel BuildCategory(PolicyCategory category)
    {
        var vm = new PolicyCategoryViewModel
        {
            Prefix = category.Prefix,
            CurrentGroups = category.CurrentGroups,
            CurrentText = category.CurrentGroups.Count == 0
                ? "(none)"
                : string.Join(", ", category.CurrentGroups.Select(g => g.DisplayName)),
            HasConflict = category.CurrentGroups.Count > 1,
        };

        foreach (var group in category.AvailableGroups)
            vm.AvailableGroups.Add(group);

        // Pre-select what the user already has, so the picker shows the current state rather than
        // an empty box that looks like "nothing assigned".
        if (category.CurrentGroups.Count == 1)
        {
            vm.SelectedGroup = vm.AvailableGroups
                .FirstOrDefault(g => string.Equals(g.Id, category.CurrentGroups[0].Id, StringComparison.OrdinalIgnoreCase));
        }

        return vm;
    }

    /// <summary>
    /// Builds the confirmation text rather than applying anything. The removal of the previous
    /// group is the part worth seeing before it happens.
    /// </summary>
    [RelayCommand]
    private void Prepare(PolicyCategoryViewModel? category)
    {
        if (category is null || _user is null)
            return;

        if (string.IsNullOrWhiteSpace(TaskNumber))
        {
            StatusMessage = "Enter the ticket/task number authorizing this change.";
            return;
        }

        var target = category.SelectedGroup;
        var removing = category.CurrentGroups
            .Where(g => target is null || !string.Equals(g.Id, target.Id, StringComparison.OrdinalIgnoreCase))
            .Select(g => g.DisplayName)
            .ToList();

        if (target is not null && removing.Count == 0 && category.CurrentGroups.Count > 0)
        {
            StatusMessage = $"{_user.DisplayName} already holds {target.DisplayName} - nothing to change.";
            return;
        }

        if (target is null && removing.Count == 0)
        {
            StatusMessage = $"{_user.DisplayName} holds no group under {category.Prefix} - nothing to change.";
            return;
        }

        category.PendingConfirmation = (target, removing.Count) switch
        {
            (null, _) => $"Remove {_user.DisplayName} from {string.Join(", ", removing)}?",
            (_, 0) => $"Add {_user.DisplayName} to {target.DisplayName}?",
            _ => $"Replace {string.Join(", ", removing)} with {target.DisplayName} for {_user.DisplayName}?",
        };
    }

    [RelayCommand]
    private void CancelPending(PolicyCategoryViewModel? category)
    {
        if (category is not null)
            category.PendingConfirmation = "";
    }

    /// <summary>
    /// Empties the picker. This is how a user is removed from a category without being given
    /// another group - an empty box means "no policy group under this prefix".
    /// </summary>
    [RelayCommand]
    private void ClearSelection(PolicyCategoryViewModel? category)
    {
        if (category is null)
            return;

        category.SelectedGroup = null;
        category.PendingConfirmation = "";
    }

    [RelayCommand]
    private async Task ConfirmAsync(PolicyCategoryViewModel? category)
    {
        if (category is null || _user is null)
            return;

        IsBusy = true;
        category.PendingConfirmation = "";
        Transcript.BeginOperation($"Assign policy group ({category.Prefix}): {_user.Upn}");

        try
        {
            var result = await _policies.AssignAsync(new AssignPolicyRequest
            {
                TaskNumber = TaskNumber.Trim(),
                User = _user,
                Group = category.SelectedGroup,
                Prefix = category.Prefix,
                Replacing = category.CurrentGroups,
            }, msg => StatusMessage = msg);

            if (!result.Succeeded)
            {
                StatusMessage = $"Change failed: {result.ErrorMessage}";
                return;
            }

            var parts = new List<string>();
            if (result.Added.Count > 0) parts.Add($"added {string.Join(", ", result.Added)}");
            if (result.Removed.Count > 0) parts.Add($"removed {string.Join(", ", result.Removed)}");

            StatusMessage = (parts.Count == 0 ? "No change was needed." : $"Done: {string.Join("; ", parts)}.")
                + (result.WarningMessage is null ? "" : $" WARNING: {result.WarningMessage}");

            // Re-read rather than patching the list locally - Graph is the truth, and a partly
            // failed removal has to show up here.
            await LookUpAsync();
        }
        catch (Exception ex)
        {
            StatusMessage = $"Change failed: {ErrorText.Describe(ex)}";
        }
        finally
        {
            IsBusy = false;
        }
    }
}
