using System.Net;

namespace M365Manager.Core.NewHire;

/// <summary>
/// Subject and HTML body for the two mails a New Hire run sends, following the wording of the legacy
/// NewHireWizardPS2.ps1. English only, like every notification this app sends - recipients span
/// countries.
/// </summary>
internal static class NewHireNotificationMail
{
    /// <summary>
    /// To the employee, once they are enabled. The 24-hour caveat is from the original and is worth
    /// keeping: the dial pad genuinely does not appear immediately.
    /// </summary>
    public static (string Subject, string Html) Enabled(string givenName, string phoneNumber)
    {
        var salutation = string.IsNullOrWhiteSpace(givenName) ? "Hello" : $"Hello {E(givenName)}";

        var html = $"""
            <p>{salutation},</p>

            <p>We have assigned telephone number <b>{E(phoneNumber)}</b> to you and enabled you for Teams
            Enterprise Voice.</p>

            <p>This allows you to make and receive telephone calls, and you now have a voicemail box. The system
            needs up to 24 hours to activate all changes - please give it until tomorrow if you do not see the
            dial pad in Teams today.</p>

            <p>Kind regards,<br>
            Unified Communications Team</p>
            """;

        return ("You have been enabled for Microsoft Teams Enterprise Voice", html);
    }

    /// <summary>
    /// To the telephony team when a run did not complete cleanly, so somebody picks up the
    /// half-configured account instead of it sitting there until the employee complains.
    /// </summary>
    public static (string Subject, string Html) Failure(
        NewHireRequest request,
        IReadOnlyList<NewHireStep> steps,
        int newHireLogId)
    {
        var failed = string.Concat(steps
            .Where(s => !s.Succeeded)
            .Select(s => $"<li><b>{E(s.Name)}</b> ({E(s.Detail)}): {E(s.Error ?? "")}</li>"));

        var completed = string.Concat(steps
            .Where(s => s.Succeeded)
            .Select(s => $"<li>{E(s.Name)}: {E(s.Detail)}</li>"));

        var html = $"""
            <p>Hello team,</p>

            <p>Enabling a new hire for telephony did not complete cleanly. The account is partly configured and
            needs to be checked by hand.</p>

            <table cellpadding="4" cellspacing="0">
              <tr><td><b>Employee</b></td><td>{E(request.User.DisplayName)}</td></tr>
              <tr><td><b>Account</b></td><td>{E(request.User.SamAccountName)} ({E(request.User.UserPrincipalName)})</td></tr>
              <tr><td><b>Office</b></td><td>{E(request.Location.Name)} ({E(request.Location.LocationCode)})</td></tr>
              <tr><td><b>Number</b></td><td>{E(request.PhoneNumber)}</td></tr>
              <tr><td><b>Ticket</b></td><td>{E(request.TaskNumber)}</td></tr>
              <tr><td><b>Run</b></td><td>newhirelog id {newHireLogId}</td></tr>
            </table>

            <p><b>Steps that failed:</b></p>
            <ul>{failed}</ul>

            <p><b>Steps that succeeded:</b></p>
            <ul>{completed}</ul>
            """;

        return ($"New hire telephony setup incomplete - {request.User.DisplayName} ({request.TaskNumber})", html);
    }

    private static string E(string value) => WebUtility.HtmlEncode(value);
}
