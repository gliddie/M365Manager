using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Controls;
using M365Manager.Core.RoomResources;

namespace M365Manager.ViewModels;

/// <summary>
/// The Access tab of the manage panel - the "Access Change to Restricted Room" ticket
/// (RoomResourceOOPChanges.ps1). The ticket names people to add or remove for two lists, the room
/// delegates and the people authorized to book; the tab has the same four fields, so a ticket is
/// copied across once and applied in one run instead of one run per list and direction.
/// </summary>
public sealed partial class RoomResourcesViewModel
{
    [ObservableProperty] private string _accessAddDelegates = "";
    [ObservableProperty] private string _accessRemoveDelegates = "";
    [ObservableProperty] private string _accessAddBookingUsers = "";
    [ObservableProperty] private string _accessRemoveBookingUsers = "";

    [ObservableProperty] private string _accessDelegatesSource = "";
    [ObservableProperty] private string _accessDelegatesMembers = "";
    [ObservableProperty] private string _accessBookingSource = "";
    [ObservableProperty] private string _accessBookingMembers = "";

    /// <summary>False on a general-use room - anyone may book it, so the booking fields are hidden.</summary>
    [ObservableProperty] private bool _hasBookingList;

    /// <summary>What is special about this room's lists (shared site-wide group, shared .DE group). Empty hides it.</summary>
    [ObservableProperty] private string _accessNote = "";

    [ObservableProperty] private bool _isAccessLoaded;
    [ObservableProperty] private string _accessLoadError = "";

    /// <summary>Guards against a slow read for the previous row landing after the next one was selected.</summary>
    private int _accessLoadVersion;

    private void ResetAccess()
    {
        _accessLoadVersion++;
        AccessAddDelegates = "";
        AccessRemoveDelegates = "";
        AccessAddBookingUsers = "";
        AccessRemoveBookingUsers = "";
        AccessDelegatesSource = "";
        AccessDelegatesMembers = "";
        AccessBookingSource = "";
        AccessBookingMembers = "";
        HasBookingList = false;
        AccessNote = "";
        AccessLoadError = "";
        IsAccessLoaded = false;
    }

    private async Task LoadAccessAsync()
    {
        if (SelectedRoom is not { } room)
            return;

        var version = ++_accessLoadVersion;
        AccessLoadError = "";

        try
        {
            var access = await _rooms.GetAccessAsync(room.PrimarySmtpAddress);
            if (version != _accessLoadVersion)
                return;

            ShowAccess(access);
        }
        catch (Exception ex)
        {
            if (version != _accessLoadVersion)
                return;

            IsAccessLoaded = false;
            AccessLoadError = $"Could not read the room's delegates and booking list: {ErrorText.Describe(ex)}";
        }
    }

    private void ShowAccess(RoomAccessSnapshot access)
    {
        static string Members(RoomAccessList list) =>
            list.Members.Count == 0 ? "(nobody)" : string.Join("\n", list.Members.Select(m => m.Label));

        AccessDelegatesSource = access.Delegates.Source;
        AccessDelegatesMembers = Members(access.Delegates);

        HasBookingList = access.BookingUsers is not null;
        AccessBookingSource = access.BookingUsers?.Source ?? "";
        AccessBookingMembers = access.BookingUsers is { } users ? Members(users) : "";

        AccessNote = access.Delegates.IsSiteWide
            ? $"{access.Delegates.Source} is the site-wide delegate group: a change to it applies to every general-use room at the site."
            : access.BookingSharesDelegateGroup
                ? "Booking rights on this room come from its delegate group, so both lists are the same group."
                : "";

        IsAccessLoaded = true;
    }

    [RelayCommand]
    private async Task ChangeAccessAsync()
    {
        if (SelectedRoom is not { } room)
        {
            ManageStatus("Select a room or resource in the grid first.", AlertSeverity.Error);
            return;
        }
        if (string.IsNullOrWhiteSpace(ManageTaskNumber))
        {
            ManageStatus("Enter the ticket/task number authorizing this change.", AlertSeverity.Error);
            return;
        }

        var request = new RoomAccessChangeRequest
        {
            TaskNumber = ManageTaskNumber.Trim(),
            RequesterIdentity = ManageRequesterIdentity.Trim(),
            Address = room.PrimarySmtpAddress,
            AddDelegates = AccessAddDelegates,
            RemoveDelegates = AccessRemoveDelegates,
            AddBookingUsers = HasBookingList ? AccessAddBookingUsers : "",
            RemoveBookingUsers = HasBookingList ? AccessRemoveBookingUsers : "",
        };
        if (request.IsEmpty)
        {
            ManageStatus("Enter at least one person to add or remove.", AlertSeverity.Error);
            return;
        }

        IsManageBusy = true;
        ManageStatus("Looking up the people in the ticket...");
        Transcript.BeginOperation($"Access change: {room.DisplayName}");
        try
        {
            var plan = await _rooms.PlanAccessChangeAsync(request);

            if (plan.Errors.Count > 0)
            {
                ManageStatus(string.Join(" ", plan.Errors) + SkippedText(plan.Skipped), AlertSeverity.Error);
                return;
            }
            if (!plan.HasWork)
            {
                ManageStatus("Nothing to change." + SkippedText(plan.Skipped), AlertSeverity.Warning);
                return;
            }

            // Names were typed from a ticket and resolved by Exchange - show who they turned out to
            // be before anything is written. Taking access away is destructive; adding is not.
            var title = $"Change access to {plan.Room.DisplayName}?";
            var message = "These changes are written to the audit log with the ticket number.";
            var detail = DescribePlan(plan);
            var confirmed = plan.HasRemovals
                ? _dialogs.ConfirmDestructive(title, message, "Remove access", detail)
                : _dialogs.Confirm(title, message, "Grant access", detail);
            if (!confirmed)
            {
                ManageStatus("Access change cancelled - nothing was changed.", AlertSeverity.Info);
                return;
            }

            ManageStatus("Applying access change...");
            var result = await _rooms.ApplyAccessChangeAsync(plan);
            ReportAccessResult(plan, result);

            if (result.Succeeded)
            {
                if (result.Results.All(r => r.Succeeded) && !plan.Skipped.Any())
                {
                    AccessAddDelegates = "";
                    AccessRemoveDelegates = "";
                    AccessAddBookingUsers = "";
                    AccessRemoveBookingUsers = "";
                }
                await LoadAccessAsync();
            }
        }
        catch (Exception ex)
        {
            ManageStatus($"Access change failed: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    private void ReportAccessResult(RoomAccessChangePlan plan, RoomAccessChangeResult result)
    {
        if (!result.Succeeded)
        {
            var reasons = result.Results.Where(r => !r.Succeeded).Select(r => $"{r.Item.Label} ({r.Error})");
            ManageStatus($"Access change failed: {result.ErrorMessage} {string.Join("; ", reasons)}".TrimEnd(), AlertSeverity.Error);
            return;
        }

        var ok = result.Results.Where(r => r.Succeeded).Select(r => r.Item).ToList();
        var failed = result.Results.Where(r => !r.Succeeded).ToList();

        var text = new StringBuilder($"Access changed for {plan.Room.DisplayName}: "
                                     + $"{ok.Count(i => i.IsAdd)} added, {ok.Count(i => !i.IsAdd)} removed.");
        if (failed.Count > 0)
            text.Append(" Failed: ").Append(string.Join("; ", failed.Select(f => $"{f.Item.Label} ({f.Error})"))).Append('.');
        text.Append(SkippedText(plan.Skipped));
        if (result.WarningMessage is not null)
            text.Append(" WARNING: ").Append(result.WarningMessage);

        var clean = failed.Count == 0 && !plan.Skipped.Any() && result.WarningMessage is null;
        ManageStatus(text.ToString(), clean ? AlertSeverity.Success : AlertSeverity.Warning);
    }

    private static string SkippedText(IEnumerable<RoomAccessChangeItem> skipped)
    {
        var list = skipped.ToList();
        return list.Count == 0
            ? ""
            : " Skipped: " + string.Join("; ", list.Select(s => $"{s.Label} ({s.Outcome.Describe()})")) + ".";
    }

    /// <summary>The confirmation's detail: per list, who is added and removed, then what is skipped and why.</summary>
    private static string DescribePlan(RoomAccessChangePlan plan)
    {
        var sb = new StringBuilder();

        void Section(RoomAccessListKind kind, string heading, string source)
        {
            var work = plan.Items.Where(i => i.List == kind && i.IsWork).ToList();
            if (work.Count == 0)
                return;

            sb.AppendLine($"{heading} ({source})");
            foreach (var item in work.Where(i => i.IsAdd))
                sb.AppendLine($"  + {item.Label}");
            foreach (var item in work.Where(i => !i.IsAdd))
                sb.AppendLine($"  − {item.Label}");
            sb.AppendLine();
        }

        Section(RoomAccessListKind.Delegates, "Room delegates", plan.Room.Delegates.Source);
        Section(RoomAccessListKind.BookingUsers, "Authorized to book", plan.Room.BookingUsers?.Source ?? "");

        var skipped = plan.Skipped.ToList();
        if (skipped.Count > 0)
        {
            sb.AppendLine("Skipped");
            foreach (var item in skipped)
                sb.AppendLine($"  {item.Label} - {item.Outcome.Describe()}");
            sb.AppendLine();
        }

        foreach (var warning in plan.Warnings)
            sb.AppendLine(warning);

        return sb.ToString().TrimEnd();
    }
}
