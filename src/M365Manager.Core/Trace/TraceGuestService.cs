using System.Text.Json;
using M365Manager.Core.M365;
using M365Manager.Core.Settings;
using M365Manager.Data.Logging;

namespace M365Manager.Core.Trace;

/// <summary>
/// See <see cref="ITraceGuestService"/>. Pure Graph - no Exchange PowerShell involved, so this
/// works whether or not the Exchange connection came up.
///
/// No new Graph consent is needed: creating an invitation and patching a user are both covered by
/// the Directory.ReadWrite.All scope the app already requests (see M365AuthService.Scopes).
/// </summary>
public sealed class TraceGuestService : ITraceGuestService
{
    private readonly GraphRestClient _graph;
    private readonly ISettingsService _settings;
    private readonly IM365AuthService _auth;
    private readonly ILogService _log;

    public TraceGuestService(GraphRestClient graph, ISettingsService settings, IM365AuthService auth, ILogService log)
    {
        _graph = graph;
        _settings = settings;
        _auth = auth;
        _log = log;
    }

    public async Task<TraceGuestInvitationResult> InviteAsync(TraceGuestInvitationRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        var trace = _settings.Current.Trace;

        try
        {
            await LogAsync(correlationId, "STAR", "Trace guest invitation has started", request, Severity.Info);

            var displayName = (request.DisplayName ?? "").Trim();
            var email = (request.Email ?? "").Trim();

            if (displayName.Length == 0)
                return await FailAsync(correlationId, request, "Enter the guest's display name.");
            if (email.Length == 0 || !email.Contains('@'))
                return await FailAsync(correlationId, request, $"'{email}' is not a valid e-mail address.");

            var redirectUrl = trace.RedirectUrl.Trim();
            if (redirectUrl.Length == 0)
                return await FailAsync(correlationId, request, "No redirect URL configured - set it under Settings > Trace application.");

            var message = BuildMessage(trace.InvitationMessage, redirectUrl, displayName);

            await LogAsync(correlationId, "INPUT",
                $"DisplayName={displayName}; Email={email}; RedirectUrl={redirectUrl}; CompanyName={trace.CompanyName}",
                request, Severity.Info);

            // 1. Create the invitation. Graph sends the mail itself.
            Progress(onProgress, $"Inviting {email}...");
            var body = new Dictionary<string, object?>
            {
                ["invitedUserDisplayName"] = displayName,
                ["invitedUserEmailAddress"] = email,
                ["inviteRedirectUrl"] = redirectUrl,
                ["sendInvitationMessage"] = true,
                ["invitedUserMessageInfo"] = new Dictionary<string, object?>
                {
                    ["customizedMessageBody"] = message,
                },
            };

            string? userId;
            string? redeemUrl;
            using (var response = await _graph.PostAsync("/invitations", body, ct))
            {
                var root = response.RootElement;
                userId = root.TryGetProperty("invitedUser", out var invitedUser)
                         && invitedUser.TryGetProperty("id", out var id)
                    ? id.GetString()
                    : null;
                redeemUrl = root.TryGetProperty("inviteRedeemUrl", out var redeem) ? redeem.GetString() : null;
            }

            await LogAsync(correlationId, "INVITE", $"Invited {displayName} <{email}> (user id {userId ?? "unknown"})", request, Severity.Success);

            // 2. Stamp the company name.
            //
            // The legacy script slept 20 seconds and then searched for the user by mail filter,
            // because New-AzureADMSInvitation didn't hand back the object. The Graph invitation
            // response carries invitedUser.id directly, so neither the wait nor the retry loop
            // that followed it is needed.
            string? warning = null;
            var companyName = trace.CompanyName.Trim();

            if (string.IsNullOrEmpty(userId))
            {
                warning = "The invitation was sent, but Graph returned no user id - set the company name manually.";
                await LogAsync(correlationId, "WARN", warning, request, Severity.Warning);
            }
            else if (companyName.Length > 0)
            {
                Progress(onProgress, "Setting the company name...");
                try
                {
                    using var _ = await _graph.PatchAsync($"/users/{userId}", new Dictionary<string, object?>
                    {
                        ["companyName"] = companyName,
                    }, ct);
                    await LogAsync(correlationId, "CHNG", $"Company name set to '{companyName}'", request, Severity.Success);
                }
                catch (Exception ex)
                {
                    // The guest exists at this point - report it, don't fail the whole operation.
                    warning = $"Guest was invited, but the company name could not be set: {ex.Message}";
                    await LogAsync(correlationId, "WARN", warning, request, Severity.Warning);
                }
            }

            await LogAsync(correlationId, "DONE", "Trace guest invitation has completed", request, Severity.Success);
            Progress(onProgress, $"{displayName} invited.");

            return new TraceGuestInvitationResult
            {
                Succeeded = true,
                DisplayName = displayName,
                Email = email,
                UserId = userId,
                RedeemUrl = redeemUrl,
                Warning = warning,
                CorrelationId = correlationId,
            };
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, request, ex.Message);
        }
    }

    /// <summary>Substitutes the placeholders the settings template may use.</summary>
    private static string BuildMessage(string template, string redirectUrl, string displayName)
        => (template ?? "")
            .Replace("{RedirectUrl}", redirectUrl)
            .Replace("{DisplayName}", displayName);

    public async Task<IReadOnlyList<TraceGuestRow>> GetGuestsAsync(CancellationToken ct = default)
    {
        var companyName = _settings.Current.Trace.CompanyName.Trim();
        if (companyName.Length == 0)
            return Array.Empty<TraceGuestRow>();

        // companyName is outside Graph's default filterable set, so this needs an advanced query
        // (ConsistencyLevel: eventual plus $count) - hence advancedQuery: true.
        var escaped = companyName.Replace("'", "''");
        var path = "/users?$filter=" + Uri.EscapeDataString($"userType eq 'Guest' and companyName eq '{escaped}'")
                   + "&$count=true&$top=100"
                   + "&$select=" + Uri.EscapeDataString("id,displayName,mail,userPrincipalName,companyName,externalUserState,externalUserStateChangeDateTime,createdDateTime,accountEnabled");

        var items = await _graph.GetAllPagesAsync(path, advancedQuery: true, ct: ct);

        return items
            .Select(item => new TraceGuestRow
            {
                Id = Str(item, "id"),
                DisplayName = Str(item, "displayName"),
                Mail = Str(item, "mail"),
                UserPrincipalName = Str(item, "userPrincipalName"),
                CompanyName = Str(item, "companyName"),
                InvitationState = Str(item, "externalUserState"),
                StateChangedUtc = Date(item, "externalUserStateChangeDateTime"),
                CreatedUtc = Date(item, "createdDateTime"),
                AccountEnabled = Bool(item, "accountEnabled"),
            })
            .OrderBy(g => g.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string Str(JsonElement element, string name)
        => element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : "";

    private static DateTime? Date(JsonElement element, string name)
        => element.TryGetProperty(name, out var value)
           && value.ValueKind == JsonValueKind.String
           && DateTime.TryParse(value.GetString(), out var parsed)
            ? parsed.ToUniversalTime()
            : null;

    private static bool Bool(JsonElement element, string name)
        => element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.True;

    // ----- logging -----

    private async Task<TraceGuestInvitationResult> FailAsync(Guid correlationId, TraceGuestInvitationRequest request, string message)
    {
        await LogAsync(correlationId, "FAIL", message, request, Severity.Error);
        return TraceGuestInvitationResult.Failed(correlationId, message);
    }

    private async Task LogAsync(Guid correlationId, string eventCode, string message, TraceGuestInvitationRequest request, Severity severity)
    {
        try
        {
            await _log.WriteAsync(new LogEntry
            {
                Area = "Trace",
                Action = "InviteGuest",
                EventCode = eventCode,
                Severity = severity,
                Message = message,
                TargetObject = request.Email,
                TaskNumber = string.IsNullOrWhiteSpace(request.TaskNumber) ? null : request.TaskNumber.Trim(),
                CorrelationId = correlationId,
                UserUpn = _auth.CurrentUser?.Upn,
            });
        }
        catch
        {
            // Logging is best-effort; never let it break the invitation.
        }
    }

    private static void Progress(Action<string>? onProgress, string message) => onProgress?.Invoke(message);
}
