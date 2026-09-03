using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using M365Manager.Core.M365;
using M365Manager.Core.Settings;

namespace M365Manager.Core.Notifications;

/// <summary>See <see cref="INotificationMailService"/>.</summary>
public sealed class NotificationMailService : INotificationMailService
{
    private static readonly char[] AddressSeparators = { ',', ';' };

    private readonly ISettingsService _settings;
    private readonly GraphRestClient _graph;

    public NotificationMailService(ISettingsService settings, GraphRestClient graph)
    {
        _settings = settings;
        _graph = graph;
    }

    public bool IsSmtpConfigured => _settings.Current.Smtp.IsConfigured;

    public string SenderDescription
    {
        get
        {
            var smtp = _settings.Current.Smtp;
            if (!smtp.IsConfigured)
                return "the signed-in account (Microsoft Graph)";

            return string.IsNullOrWhiteSpace(smtp.FromDisplayName)
                ? smtp.FromAddress
                : $"{smtp.FromDisplayName} <{smtp.FromAddress}>";
        }
    }

    public Task SendAsync(
        IReadOnlyList<string> recipients,
        string subject,
        string htmlBody,
        IReadOnlyList<InlineImage>? inlineImages = null,
        CancellationToken ct = default)
    {
        var to = recipients
            .Where(r => !string.IsNullOrWhiteSpace(r))
            .Select(r => r.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (to.Length == 0)
            throw new InvalidOperationException("There is no recipient address to send the notification to.");

        var images = inlineImages ?? Array.Empty<InlineImage>();
        var smtp = _settings.Current.Smtp;
        return smtp.IsConfigured
            ? SendViaSmtpAsync(smtp, to, subject, htmlBody, images, ct)
            : SendViaGraphAsync(to, subject, htmlBody, images, ct);
    }

    private static async Task SendViaSmtpAsync(SmtpSettings smtp, string[] to, string subject, string htmlBody, IReadOnlyList<InlineImage> images, CancellationToken ct)
    {
        using var message = new MailMessage
        {
            From = new MailAddress(
                smtp.FromAddress.Trim(),
                string.IsNullOrWhiteSpace(smtp.FromDisplayName) ? smtp.FromAddress.Trim() : smtp.FromDisplayName.Trim()),
            Subject = subject,
        };

        // With inline images the body has to become an AlternateView so the images can hang off it
        // as linked resources; without them a plain HTML body is enough.
        if (images.Count == 0)
        {
            message.Body = htmlBody;
            message.IsBodyHtml = true;
        }
        else
        {
            var view = AlternateView.CreateAlternateViewFromString(htmlBody, null, MediaTypeNames.Text.Html);
            foreach (var image in images)
            {
                var resource = new LinkedResource(new MemoryStream(image.Content), image.MediaType)
                {
                    ContentId = image.ContentId,
                    // Inline, not an attachment the recipient sees listed separately.
                    TransferEncoding = TransferEncoding.Base64,
                };
                resource.ContentLink = new Uri($"cid:{image.ContentId}");
                view.LinkedResources.Add(resource);
            }
            message.AlternateViews.Add(view);
        }

        foreach (var recipient in to)
            message.To.Add(recipient);

        if (!string.IsNullOrWhiteSpace(smtp.ReplyToAddress))
            message.ReplyToList.Add(smtp.ReplyToAddress.Trim());

        foreach (var bcc in SplitAddresses(smtp.BccAddress))
            message.Bcc.Add(bcc);

        using var client = new SmtpClient(smtp.Host.Trim(), smtp.Port > 0 ? smtp.Port : 25)
        {
            EnableSsl = smtp.UseSsl,
        };

        // Order matters: assigning UseDefaultCredentials resets Credentials, so it goes first.
        if (!string.IsNullOrWhiteSpace(smtp.UserName))
        {
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential(smtp.UserName.Trim(), smtp.Password);
        }

        await client.SendMailAsync(message, ct);
    }

    private async Task SendViaGraphAsync(string[] to, string subject, string htmlBody, IReadOnlyList<InlineImage> images, CancellationToken ct)
    {
        var message = new Dictionary<string, object?>
        {
            ["subject"] = subject,
            ["body"] = new Dictionary<string, object?> { ["contentType"] = "HTML", ["content"] = htmlBody },
            ["toRecipients"] = to.Select(address => new Dictionary<string, object?>
            {
                ["emailAddress"] = new Dictionary<string, object?> { ["address"] = address },
            }).ToArray(),
        };

        if (images.Count > 0)
        {
            // isInline + contentId is what makes Graph render these in the body rather than
            // listing them as attachments - the equivalent of SMTP's linked resources above.
            message["attachments"] = images.Select(image => new Dictionary<string, object?>
            {
                ["@odata.type"] = "#microsoft.graph.fileAttachment",
                ["name"] = image.FileName,
                ["contentType"] = image.MediaType,
                ["contentId"] = image.ContentId,
                ["isInline"] = true,
                ["contentBytes"] = Convert.ToBase64String(image.Content),
            }).ToArray();
        }

        var body = new Dictionary<string, object?>
        {
            ["message"] = message,
            ["saveToSentItems"] = true,
        };

        using var _ = await _graph.PostAsync("/me/sendMail", body, ct);
    }

    private static IEnumerable<string> SplitAddresses(string value) =>
        string.IsNullOrWhiteSpace(value)
            ? Array.Empty<string>()
            : value.Split(AddressSeparators, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
}
