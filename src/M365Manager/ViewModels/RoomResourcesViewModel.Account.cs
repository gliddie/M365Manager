using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using M365Manager.Controls;

namespace M365Manager.ViewModels;

/// <summary>
/// The Account tab of the manage panel: enables the selected room's account with a new password
/// that never expires - for a room whose account was never enabled, or a Teams Rooms device that
/// needs a new password. Same copy card as after Create.
/// </summary>
public sealed partial class RoomResourcesViewModel
{
    [ObservableProperty] private string _managePassword = "";
    [ObservableProperty] private string _managePasswordUpn = "";
    [ObservableProperty] private string _managePasswordDisplayName = "";

    public bool HasManagePassword => ManagePassword.Length > 0;

    partial void OnManagePasswordChanged(string value) => OnPropertyChanged(nameof(HasManagePassword));

    /// <summary>Called when the grid selection changes - a password must never sit next to another room.</summary>
    private void ClearManagePassword()
    {
        ManagePassword = "";
        ManagePasswordUpn = "";
        ManagePasswordDisplayName = "";
    }

    [RelayCommand]
    private async Task SetNewAccountPasswordAsync()
    {
        if (SelectedRoom is not { } room)
        {
            ManageStatus("Select a room in the grid first.", AlertSeverity.Error);
            return;
        }

        if (string.Equals(room.ResourceKind, "Equipment", StringComparison.OrdinalIgnoreCase))
        {
            ManageStatus("Only room accounts are enabled for sign-in - equipment mailboxes have no device that signs in.", AlertSeverity.Error);
            return;
        }

        if (string.IsNullOrWhiteSpace(ManageTaskNumber))
        {
            ManageStatus("Enter the ticket/task number authorizing this change.", AlertSeverity.Error);
            return;
        }

        // A new password signs out every device using the old one - so it is confirmed like any
        // other action that breaks something in use.
        if (!_dialogs.ConfirmDestructive(
                $"Set a new password for {room.DisplayName}?",
                "The room account is enabled with a new random password that never expires. "
                + "A Teams Rooms device signed in with the current password is signed out and has to be signed in again.",
                "Set new password",
                $"{room.PrimarySmtpAddress}\n\nThe password is shown once here to copy. It is not stored, logged or mailed."))
            return;

        IsManageBusy = true;
        ClearManagePassword();
        ManageStatus("Setting a new password...");
        Transcript.BeginOperation($"Room account password: {room.PrimarySmtpAddress}");
        try
        {
            var result = await _rooms.SetNewAccountPasswordAsync(room.PrimarySmtpAddress, ManageTaskNumber.Trim());
            if (!ReferenceEquals(room, SelectedRoom))
                return; // never show a password next to a different room

            if (!result.Succeeded)
            {
                ManageStatus($"The password could not be set: {result.ErrorMessage}", AlertSeverity.Error);
                return;
            }

            ManagePassword = result.Password;
            ManagePasswordUpn = result.UserPrincipalName;
            ManagePasswordDisplayName = result.DisplayName.Length > 0 ? result.DisplayName : room.DisplayName;
            ManageStatus(result.WarningMessage is null
                    ? $"{room.DisplayName}: account enabled with a new password that never expires. Copy it from the Account tab."
                    : $"{room.DisplayName}: account enabled with a new password. WARNING: {result.WarningMessage}",
                result.WarningMessage is null ? AlertSeverity.Success : AlertSeverity.Warning);
        }
        catch (Exception ex)
        {
            ManageStatus($"The password could not be set: {ErrorText.Describe(ex)}", AlertSeverity.Error);
        }
        finally
        {
            IsManageBusy = false;
        }
    }

    [RelayCommand]
    private void CopyManagePassword() => CopyToManageClipboard(ManagePassword, "Password");

    [RelayCommand]
    private void CopyManageUpn() => CopyToManageClipboard(ManagePasswordUpn, "Sign-in name");

    [RelayCommand]
    private void CopyManageDisplayName() => CopyToManageClipboard(ManagePasswordDisplayName, "Display name");

    /// <summary>Like the Create tab's copy, but reporting in the manage panel's status.</summary>
    private void CopyToManageClipboard(string value, string what)
    {
        if (value.Length == 0)
            return;

        try
        {
            System.Windows.Clipboard.SetText(value);
            ManageStatus($"{what} copied to the clipboard.", AlertSeverity.Success);
        }
        catch (Exception ex)
        {
            ManageStatus($"Could not copy - select the text and press Ctrl+C instead ({ex.Message}).", AlertSeverity.Warning);
        }
    }
}
