namespace M365Manager.Core.Notifications;

/// <summary>
/// Sends the notification e-mails the admin features produce (new shared mailbox, owner change,
/// new team, ...). Two transports, picked per call:
///
/// 1. The SMTP relay from <see cref="M365Manager.Core.Settings.SmtpSettings"/>, when configured.
///    Preferred - the mail comes from a fixed team address and works with an admin account that
///    has no Exchange mailbox.
/// 2. Graph /me/sendMail as the signed-in admin, when no relay is configured. This is what the
///    app did before SMTP existed; it fails for unlicensed admin accounts.
/// </summary>
public interface INotificationMailService
{
    /// <summary>True when an SMTP relay is configured, i.e. mail will NOT go out as the signed-in user.</summary>
    bool IsSmtpConfigured { get; }

    /// <summary>Human-readable sender, for progress text and log lines ("Team &lt;team@x&gt;" / "the signed-in account").</summary>
    string SenderDescription { get; }

    /// <summary>
    /// Sends an HTML mail to every recipient. Throws on failure - callers decide whether a failed
    /// notification should abort their operation (it generally shouldn't) but must surface it.
    /// </summary>
    /// <param name="inlineImages">
    /// Screenshots/logos the body references as &lt;img src="cid:..."&gt;, as the legacy .oft
    /// templates do. Each image's <see cref="InlineImage.ContentId"/> has to match the cid in the
    /// body exactly. Empty for a plain-text-style notification.
    /// </param>
    Task SendAsync(
        IReadOnlyList<string> recipients,
        string subject,
        string htmlBody,
        IReadOnlyList<InlineImage>? inlineImages = null,
        CancellationToken ct = default);
}

/// <summary>
/// An image embedded in the mail itself rather than linked to a web server - the only way a
/// screenshot in a notification renders reliably, since mail clients block remote images by default.
/// </summary>
/// <param name="ContentId">Matches the "cid:" reference in the HTML body, without the "cid:" prefix.</param>
/// <param name="FileName">Shown by clients that list inline parts as attachments.</param>
/// <param name="MediaType">e.g. "image/png".</param>
/// <param name="Content">The raw image bytes.</param>
public sealed record InlineImage(string ContentId, string FileName, string MediaType, byte[] Content);
