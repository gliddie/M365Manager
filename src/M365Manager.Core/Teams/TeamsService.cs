using System.Data;
using System.Management.Automation;
using System.Text;
using System.Text.Json;
using M365Manager.Core.M365;
using M365Manager.Core.Notifications;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Core.SharePoint;
using M365Manager.Data.Logging;
using Microsoft.Data.SqlClient;

namespace M365Manager.Core.Teams;

/// <summary>
/// See <see cref="ITeamsService"/>. Graph calls that this codebase hasn't used before
/// (group creation with owners@odata.bind, teamify, guest directory settings, sendMail) go
/// through plain REST via <see cref="IM365AuthService.GetAccessTokenAsync"/> rather than the
/// Kiota SDK client, so the exact request/response JSON shape (stable, documented Graph v1.0
/// contracts) isn't at the mercy of a particular SDK version's generated type surface. User
/// lookup reuses <see cref="IM365AuthService.Graph"/> - the same call shape already proven
/// elsewhere in this codebase (see ExchangeService.RemoveMembersAsync).
/// </summary>
public sealed class TeamsService : ITeamsService
{
    /// <summary>Only still needed to build the absolute URLs an @odata.bind reference requires.</summary>
    private const string GraphBaseUrl = "https://graph.microsoft.com/v1.0";

    private readonly PowerShellHost _host;
    private readonly IM365AuthService _auth;
    private readonly ISettingsService _settings;
    private readonly ITeamNamingService _naming;
    private readonly ISharePointService _sharePoint;
    private readonly ILogService _log;
    private readonly IConnectionStringProvider _connectionStrings;
    private readonly INotificationMailService _mail;
    private readonly GraphRestClient _graph;

    public TeamsService(
        PowerShellHost host,
        IM365AuthService auth,
        ISettingsService settings,
        ITeamNamingService naming,
        ISharePointService sharePoint,
        ILogService log,
        IConnectionStringProvider connectionStrings,
        INotificationMailService mail,
        GraphRestClient graph)
    {
        _graph = graph;
        _host = host;
        _auth = auth;
        _settings = settings;
        _naming = naming;
        _sharePoint = sharePoint;
        _log = log;
        _connectionStrings = connectionStrings;
        _mail = mail;
    }

    public async Task<TeamCreationResult> CreateTeamAsync(TeamCreationRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        var report = (string code, string message) => LogAsync(correlationId, code, message, request.TaskNumber, Severity.Info);

        try
        {
            await report("STAR", "New Team script has started");
            Progress(onProgress, "Starting team creation...");

            // 1. Owner resolution + licensing check.
            var owner = await ResolveOwnerAsync(request.OwnerIdentity, ct);
            if (owner is null)
                return await FailAsync(correlationId, request.TaskNumber, $"Could not resolve owner '{request.OwnerIdentity}' to a directory user.");

            Progress(onProgress, $"Owner resolved: {owner.Value.Upn}");
            var hasLicensedMailbox = await OwnerHasLicensedMailboxAsync(owner.Value.Upn, ct);
            if (!hasLicensedMailbox)
                return await FailAsync(correlationId, request.TaskNumber, $"Owner '{owner.Value.Upn}' does not have a licensed mailbox - creation cancelled.");

            // 2. Build the enforced name + alias.
            var displayName = await _naming.BuildDisplayNameAsync(request.Location, request.Name, ct);
            var alias = _naming.BuildAlias(displayName);
            Progress(onProgress, $"Team name: {displayName} (alias: {alias})");

            // 3. Collision check - but a group left behind by an earlier failed run (e.g. the
            //    teamify-timing issue handled in step 5) isn't a hard collision: resume
            //    configuring it instead of creating a duplicate or blocking forever.
            await LogAsync(correlationId, "INPUT", $"Name={displayName}; Alias={alias}; Owner={owner.Value.Upn}; Public={request.IsPublic}; Internal={request.IsInternal}", request.TaskNumber, Severity.Info);

            string groupId;
            var existingGroup = await GetUnifiedGroupAsync(alias, ct);
            if (existingGroup is not null)
            {
                var existingGroupId = Str(existingGroup, "ExternalDirectoryObjectId");
                if (string.IsNullOrWhiteSpace(existingGroupId))
                    return await FailAsync(correlationId, request.TaskNumber, $"A group named '{displayName}' (alias '{alias}') already exists, but its directory id could not be read to resume configuration.");

                if (await IsGroupATeamAsync(existingGroupId, ct))
                    return await FailAsync(correlationId, request.TaskNumber, $"A team named '{displayName}' (alias '{alias}') already exists.");

                groupId = existingGroupId;
                Progress(onProgress, $"Found existing group '{displayName}' without a Team yet - resuming from there.");
                await LogAsync(correlationId, "RESUME", $"Found existing group {displayName} ({alias}) without a Team - resuming configuration.", request.TaskNumber, Severity.Info);
            }
            else
            {
                // 4. Create the M365 group with the resolved owner set directly - the signed-in
                //    admin never becomes a member (avoids the legacy script's self-removal step).
                Progress(onProgress, "Creating Microsoft 365 group...");
                groupId = await CreateGroupAsync(displayName, alias, request.Description, request.IsPublic, owner.Value.Id, ct);
                await LogAsync(correlationId, "CREATE", $"{displayName} ({alias})", request.TaskNumber, Severity.Success);
            }

            // 5. Teamify the group.
            Progress(onProgress, "Enabling Microsoft Teams on the group...");
            await TeamifyGroupAsync(groupId, onProgress, ct);

            // 6. Wait for Exchange Online to see the new UnifiedGroup (Create-Complete equivalent).
            Progress(onProgress, "Waiting for group provisioning to complete...");
            var group = await WaitForUnifiedGroupAsync(alias, onProgress, ct);
            if (group is null)
                return await FailAsync(correlationId, request.TaskNumber, "Timed out waiting for the group to appear in Exchange Online.");

            // 7. Exchange-side settings: hide from GAL, disable welcome message, set Notes.
            Progress(onProgress, "Hiding from address book and disabling welcome message...");
            var notes = $"{request.Description}\nRequested by: {owner.Value.DisplayName}\nPer: {request.TaskNumber}";
            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-UnifiedGroup")
                .AddParameter("Identity", alias)
                .AddParameter("HiddenFromAddressListsEnabled", true)
                .AddParameter("UnifiedGroupWelcomeMessageEnabled", false)
                .AddParameter("Notes", notes)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            await LogAsync(correlationId, "DISAB", "Disabling Welcome Message", request.TaskNumber, Severity.Success);
            await LogAsync(correlationId, "HIDE", "Hiding from Address Book", request.TaskNumber, Severity.Success);

            // 8. Guest access directory setting.
            Progress(onProgress, "Setting guest access policy...");
            try
            {
                await SetGuestAccessAsync(groupId, allowGuests: !request.IsInternal, ct);
                await LogAsync(correlationId, "CHNG", $"Guest access {(request.IsInternal ? "disabled" : "enabled")}", request.TaskNumber, Severity.Success);
            }
            catch (Exception ex)
            {
                await LogAsync(correlationId, "WARN", $"Could not set guest access policy: {ex.Message}", request.TaskNumber, Severity.Warning);
            }

            // 9. SharePoint sharing lockdown for internal teams.
            if (request.IsInternal)
            {
                var siteUrl = Str(group, "SharePointSiteUrl");
                if (string.IsNullOrWhiteSpace(siteUrl))
                {
                    Progress(onProgress, "Waiting for SharePoint site provisioning...");
                    siteUrl = await WaitForSharePointSiteUrlAsync(alias, onProgress, ct);
                }

                if (string.IsNullOrWhiteSpace(siteUrl))
                {
                    await LogAsync(correlationId, "FAIL", "SharePoint site did not provision in time - sharing was not locked down.", request.TaskNumber, Severity.Warning);
                }
                else
                {
                    Progress(onProgress, "Locking down SharePoint site sharing...");
                    try
                    {
                        if (!_sharePoint.IsConnected)
                            await _sharePoint.ConnectAsync(onProgress, ct);
                        await _sharePoint.LockSiteSharingAsync(siteUrl, ct);
                        await LogAsync(correlationId, "CHNG", "Disabled SharePoint site sharing capability", request.TaskNumber, Severity.Success);
                    }
                    catch (Exception ex)
                    {
                        await LogAsync(correlationId, "FAIL", $"Could not lock SharePoint site sharing: {ex.Message}", request.TaskNumber, Severity.Warning);
                    }
                }
            }

            // 10. Owner notification mail (best-effort, but the failure is reported - see
            // SharedMailboxService for the same pattern).
            Progress(onProgress, $"Sending confirmation e-mail to the owner as {_mail.SenderDescription}...");
            try
            {
                // The mail attribute, not the UPN - see ResolveOwnerAsync.
                var ownerMail = string.IsNullOrWhiteSpace(owner.Value.Mail) ? owner.Value.Upn : owner.Value.Mail;
                await SendOwnerMailAsync(ownerMail, owner.Value.DisplayName, displayName, request.TaskNumber, ct);
                await LogAsync(correlationId, "SEND", $"Sent confirmation mail via {_mail.SenderDescription} to {ownerMail}", request.TaskNumber, Severity.Success);
            }
            catch (Exception ex)
            {
                await LogAsync(correlationId, "WARN", $"Could not send confirmation mail: {ex.Message}", request.TaskNumber, Severity.Warning);
                Progress(onProgress, $"WARNING: The team was created, but the confirmation e-mail to {owner.Value.DisplayName} could not be sent: {ex.Message}");
            }

            await LogAsync(correlationId, "DONE", "New Team script has completed", request.TaskNumber, Severity.Success);
            Progress(onProgress, $"Team '{displayName}' created successfully.");

            // Keep the SQL cache current immediately - don't make the grid wait for the next
            // scheduled import to show a team created just now (see GetTeamsOverviewAsync).
            await UpsertCacheAsync(request, groupId, displayName, alias, group, owner.Value.Upn, ct);

            return new TeamCreationResult
            {
                Succeeded = true,
                DisplayName = displayName,
                PrimarySmtpAddress = Str(group, "PrimarySmtpAddress"),
                GroupId = groupId,
                ResolvedOwnerUpn = owner.Value.Upn,
                CorrelationId = correlationId,
            };
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    // ----- steps -----

    /// <summary>
    /// <see cref="Mail"/> is the directory's mail attribute, empty when there is none. It is kept
    /// separate from the UPN because the two live on different domains in this tenant (sign-in vs.
    /// mail) - the notification e-mail needs the deliverable one, everything else the UPN.
    /// </summary>
    private async Task<(string Id, string Upn, string DisplayName, string Mail)?> ResolveOwnerAsync(string identity, CancellationToken ct)
    {
        var value = (identity ?? "").Trim();
        if (value.Length == 0)
            return null;

        // Bare SamAccountName -> complete with the configured default domain, same rule as
        // ExchangeService.ResolveMemberIdentity.
        var candidate = value.Contains('@')
            ? value
            : AppendDomain(value);

        try
        {
            using var json = await _graph.GetAsync($"/users/{Uri.EscapeDataString(candidate)}?{UserSelect}", ct: ct);
            var user = json.RootElement;
            var id = GraphText(user, "id");
            if (id.Length > 0)
                return (id,
                    Or(GraphText(user, "userPrincipalName"), candidate),
                    Or(GraphText(user, "displayName"), candidate),
                    GraphText(user, "mail"));
        }
        catch
        {
            // fall through to mail-filter lookup below
        }

        if (value.Contains('@'))
        {
            var filter = Uri.EscapeDataString($"mail eq '{value.Replace("'", "''")}'");
            using var json = await _graph.GetAsync($"/users?$filter={filter}&{UserSelect}", ct: ct);

            if (json.RootElement.TryGetProperty("value", out var matches) && matches.GetArrayLength() > 0)
            {
                var match = matches[0];
                var id = GraphText(match, "id");
                if (id.Length > 0)
                    return (id,
                        Or(GraphText(match, "userPrincipalName"), value),
                        Or(GraphText(match, "displayName"), value),
                        Or(GraphText(match, "mail"), value));
            }
        }

        return null;
    }

    private string AppendDomain(string samAccountName)
    {
        var domain = _settings.Current.M365.Domain.Trim();
        return domain.Length == 0 ? samAccountName : $"{samAccountName}@{domain}";
    }

    private async Task<bool> OwnerHasLicensedMailboxAsync(string upn, CancellationToken ct)
    {
        var results = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-EXOMailbox")
            .AddParameter("Identity", upn)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);
        return results.Count > 0;
    }

    private async Task<PSObject?> GetUnifiedGroupAsync(string alias, CancellationToken ct)
    {
        var results = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-UnifiedGroup")
            .AddParameter("Identity", alias)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);
        return results.Count > 0 ? results[0] : null;
    }

    /// <summary>Distinguishes a plain M365 group (not yet teamified) from an existing Team.</summary>
    private async Task<bool> IsGroupATeamAsync(string groupId, CancellationToken ct)
    {
        try
        {
            using var _ = await _graph.GetAsync($"/groups/{groupId}/team", ct: ct);
            return true;
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("(404)"))
        {
            return false;
        }
    }

    private async Task<PSObject?> WaitForUnifiedGroupAsync(string alias, Action<string>? onProgress, CancellationToken ct)
    {
        for (var attempt = 0; attempt < 20; attempt++)
        {
            var group = await GetUnifiedGroupAsync(alias, ct);
            if (group is not null)
                return group;

            Progress(onProgress, $"Still waiting for group provisioning... ({attempt + 1}/20)");
            await Task.Delay(TimeSpan.FromSeconds(15), ct);
        }
        return null;
    }

    private async Task<string?> WaitForSharePointSiteUrlAsync(string alias, Action<string>? onProgress, CancellationToken ct)
    {
        for (var attempt = 0; attempt < 20; attempt++)
        {
            var group = await GetUnifiedGroupAsync(alias, ct);
            var url = group is not null ? Str(group, "SharePointSiteUrl") : "";
            if (!string.IsNullOrWhiteSpace(url))
                return url;

            Progress(onProgress, $"Still waiting for SharePoint site provisioning... ({attempt + 1}/20)");
            await Task.Delay(TimeSpan.FromSeconds(15), ct);
        }
        return null;
    }

    // ----- Graph REST helpers (plain JSON - see class remarks) -----

    private async Task<string> CreateGroupAsync(string displayName, string alias, string description, bool isPublic, string ownerId, CancellationToken ct)
    {
        var body = new Dictionary<string, object?>
        {
            ["displayName"] = displayName,
            ["mailNickname"] = alias,
            ["description"] = description,
            ["mailEnabled"] = true,
            ["securityEnabled"] = false,
            ["groupTypes"] = new[] { "Unified" },
            ["visibility"] = isPublic ? "Public" : "Private",
            ["owners@odata.bind"] = new[] { $"{GraphBaseUrl}/users/{ownerId}" },
            ["members@odata.bind"] = new[] { $"{GraphBaseUrl}/users/{ownerId}" },
        };

        using var json = await _graph.PostAsync("/groups", body, ct);
        return json.RootElement.GetProperty("id").GetString()!;
    }

    private async Task TeamifyGroupAsync(string groupId, Action<string>? onProgress, CancellationToken ct)
    {
        // A group just created via POST /groups isn't always visible yet to Teams' own
        // provisioning backend - PUT /groups/{id}/team can fail with a 404
        // "GetGroupInternalApiRequest ... Not Found" for the first tens of seconds. Retry with
        // backoff instead of failing immediately, mirroring the legacy script's Create-Complete
        // wait loop for group propagation.
        const int maxAttempts = 10;
        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                // Empty body: default Team settings are fine, matching the legacy script (which
                // didn't customize Teams-specific settings either).
                using var _ = await _graph.PutAsync($"/groups/{groupId}/team", new Dictionary<string, object?>(), ct);
                return;
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("(404)") && attempt < maxAttempts)
            {
                Progress(onProgress, $"Group not yet visible to Teams provisioning, retrying... ({attempt}/{maxAttempts})");
                await Task.Delay(TimeSpan.FromSeconds(15), ct);
            }
        }
    }

    private async Task SetGuestAccessAsync(string groupId, bool allowGuests, CancellationToken ct)
    {
        using var templatesJson = await _graph.GetAsync("/groupSettingTemplates", ct: ct);

        // FirstOrDefault over JsonElement yields Undefined rather than null when nothing matches,
        // so chaining GetProperty straight onto it throws something unreadable. Find it explicitly
        // and say what actually went wrong.
        string? templateId = null;
        foreach (var template in templatesJson.RootElement.GetProperty("value").EnumerateArray())
        {
            if (template.TryGetProperty("displayName", out var displayName)
                && displayName.GetString() == "Group.Unified.Guest"
                && template.TryGetProperty("id", out var id))
            {
                templateId = id.GetString();
                break;
            }
        }

        if (templateId is null)
            throw new InvalidOperationException(
                "The 'Group.Unified.Guest' group setting template was not found in this tenant - guest access was left at the tenant default.");

        var body = new Dictionary<string, object?>
        {
            ["templateId"] = templateId,
            ["values"] = new[]
            {
                new Dictionary<string, object?> { ["name"] = "AllowToAddGuests", ["value"] = allowGuests ? "true" : "false" },
            },
        };

        using var _ = await _graph.PostAsync($"/groups/{groupId}/settings", body, ct);
    }

    /// <summary>
    /// English only, like every notification this app sends - the recipients are spread across
    /// countries. There is no legacy .oft template for team creation, so this wording follows the
    /// shared mailbox ones (see SharedMailboxService) rather than being derived from one.
    /// </summary>
    private async Task SendOwnerMailAsync(string ownerAddress, string ownerDisplayName, string teamName, string taskNumber, CancellationToken ct)
    {
        var html = $"""
            <p>Per the subject service desk ticket the team <b>{WebUtility_HtmlEscape(teamName)}</b> has been
            created for you. You are an owner of this team and are therefore responsible for managing its
            membership.</p>

            <p>You will find the team in the left-hand navigation pane of the Teams app. Using the "..." menu
            next to the team name you can open "Manage team", where you can add members and further owners.</p>

            <p>It may take up to 72 hours for the team and its address to sync to the Offline Address List.
            In the meanwhile you can find it by selecting the Global Address List.</p>

            <p>In the future if this team is no longer needed/used please create a request with the Service
            Desk so it can be removed from the system.</p>

            <p>With this information this ticket is being closed.</p>
            """;

        await _mail.SendAsync(new[] { ownerAddress }, $"{teamName} Team - {taskNumber}", html, ct: ct);
    }

    private static string WebUtility_HtmlEscape(string value) => System.Net.WebUtility.HtmlEncode(value);

    private const string UserSelect = "$select=id,userPrincipalName,displayName,mail";

    private static string GraphText(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : "";

    private static string Or(string value, string fallback) => value.Length > 0 ? value : fallback;

    // ----- SQL cache (read for the overview grid, write-through after creation) -----

    /// <summary>
    /// Reads the Teams overview grid straight from SQL (dbo.Teams JOIN dbo.Groups, member count
    /// from dbo.GroupMembers) - never live from M365. Populated by the scheduled
    /// Import-TeamsToSql.ps1 script, plus <see cref="UpsertCacheAsync"/> for teams created here.
    /// </summary>
    public async Task<IReadOnlyList<TeamOverviewRow>> GetTeamsOverviewAsync(CancellationToken ct = default)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return Array.Empty<TeamOverviewRow>();

        const string query = """
            SELECT
                g.Id, g.DisplayName, g.Alias, g.PrimarySmtpAddress, g.ManagedBy,
                t.Visibility, t.IsArchived, t.AllowToAddGuests, t.HiddenFromAddressListsEnabled,
                t.WelcomeMessageEnabled, t.SharePointSiteUrl, g.IsDeletedInM365,
                t.CreatedDateTime, t.LastImportedAtUtc,
                (SELECT COUNT(*) FROM dbo.GroupMembers gm WHERE gm.GroupId = g.Id AND gm.IsNested = 0) AS MemberCount
            FROM dbo.Teams t
            JOIN dbo.Groups g ON g.Id = t.GroupId
            ORDER BY g.DisplayName;
            """;

        await using var conn = new SqlConnection(cs);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(query, conn);
        await using var reader = await cmd.ExecuteReaderAsync(ct);

        var list = new List<TeamOverviewRow>();
        while (await reader.ReadAsync(ct))
        {
            list.Add(new TeamOverviewRow
            {
                GroupId = reader.GetGuid(0).ToString(),
                DisplayName = reader.IsDBNull(1) ? "" : reader.GetString(1),
                Alias = reader.IsDBNull(2) ? "" : reader.GetString(2),
                PrimarySmtpAddress = reader.IsDBNull(3) ? "" : reader.GetString(3),
                Owners = reader.IsDBNull(4) ? "" : reader.GetString(4).Replace(';', ',').Replace(",", ", "),
                Visibility = reader.IsDBNull(5) ? "" : reader.GetString(5),
                IsArchived = !reader.IsDBNull(6) && reader.GetBoolean(6),
                AllowToAddGuests = reader.IsDBNull(7) ? null : reader.GetBoolean(7),
                HiddenFromAddressListsEnabled = reader.IsDBNull(8) ? null : reader.GetBoolean(8),
                WelcomeMessageEnabled = reader.IsDBNull(9) ? null : reader.GetBoolean(9),
                SharePointSiteUrl = reader.IsDBNull(10) ? "" : reader.GetString(10),
                IsDeletedInM365 = !reader.IsDBNull(11) && reader.GetBoolean(11),
                CreatedDateTime = reader.IsDBNull(12) ? null : reader.GetDateTime(12),
                LastImportedAtUtc = reader.IsDBNull(13) ? default : reader.GetDateTime(13),
                MemberCount = reader.IsDBNull(14) ? 0 : reader.GetInt32(14),
            });
        }
        return list;
    }

    /// <summary>
    /// Writes the just-created group/team straight into dbo.Groups + dbo.Teams so the overview
    /// grid shows it immediately, without waiting for the next scheduled import run. Best-effort:
    /// the scheduled import reconciles this row anyway on its next pass.
    /// </summary>
    private async Task UpsertCacheAsync(TeamCreationRequest request, string groupId, string displayName, string alias, PSObject? group, string ownerUpn, CancellationToken ct)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return;

        try
        {
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync(ct);

            const string groupSql = """
                MERGE dbo.Groups AS target
                USING (SELECT @Id AS Id) AS source ON target.Id = source.Id
                WHEN MATCHED THEN UPDATE SET
                    DisplayName = @DisplayName, MailNickname = @Alias, Alias = @Alias,
                    PrimarySmtpAddress = @PrimarySmtpAddress, GroupTypes = 'Unified',
                    MailEnabled = 1, SecurityEnabled = 0, ManagedBy = @Owner,
                    UpdatedDateTime = SYSUTCDATETIME(), LastImportedAtUtc = SYSUTCDATETIME(),
                    IsDeletedInM365 = 0, LastSeenAtUtc = SYSUTCDATETIME()
                WHEN NOT MATCHED THEN INSERT
                    (Id, DisplayName, MailNickname, Alias, PrimarySmtpAddress, GroupTypes, MailEnabled,
                     SecurityEnabled, ManagedBy, CreatedDateTime, UpdatedDateTime, LastImportedAtUtc,
                     IsDeletedInM365, LastSeenAtUtc)
                VALUES
                    (@Id, @DisplayName, @Alias, @Alias, @PrimarySmtpAddress, 'Unified', 1, 0, @Owner,
                     SYSUTCDATETIME(), SYSUTCDATETIME(), SYSUTCDATETIME(), 0, SYSUTCDATETIME());
                """;

            await using (var cmd = new SqlCommand(groupSql, conn))
            {
                cmd.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = Guid.Parse(groupId);
                cmd.Parameters.Add("@DisplayName", SqlDbType.NVarChar, 256).Value = displayName;
                cmd.Parameters.Add("@Alias", SqlDbType.NVarChar, 128).Value = alias;
                cmd.Parameters.Add("@PrimarySmtpAddress", SqlDbType.NVarChar, 256).Value = NullableParam(Str(group, "PrimarySmtpAddress"));
                cmd.Parameters.Add("@Owner", SqlDbType.NVarChar, -1).Value = ownerUpn;
                await cmd.ExecuteNonQueryAsync(ct);
            }

            const string teamSql = """
                MERGE dbo.Teams AS target
                USING (SELECT @GroupId AS GroupId) AS source ON target.GroupId = source.GroupId
                WHEN MATCHED THEN UPDATE SET
                    Visibility = @Visibility, IsArchived = 0, AllowToAddGuests = @AllowToAddGuests,
                    HiddenFromAddressListsEnabled = 1, WelcomeMessageEnabled = 0,
                    SharePointSiteUrl = @SiteUrl, LastImportedAtUtc = SYSUTCDATETIME()
                WHEN NOT MATCHED THEN INSERT
                    (GroupId, Visibility, IsArchived, AllowToAddGuests, HiddenFromAddressListsEnabled,
                     WelcomeMessageEnabled, SharePointSiteUrl, CreatedDateTime, LastImportedAtUtc)
                VALUES
                    (@GroupId, @Visibility, 0, @AllowToAddGuests, 1, 0, @SiteUrl, SYSUTCDATETIME(), SYSUTCDATETIME());
                """;

            await using (var cmd = new SqlCommand(teamSql, conn))
            {
                cmd.Parameters.Add("@GroupId", SqlDbType.UniqueIdentifier).Value = Guid.Parse(groupId);
                cmd.Parameters.Add("@Visibility", SqlDbType.NVarChar, 32).Value = request.IsPublic ? "Public" : "Private";
                cmd.Parameters.Add("@AllowToAddGuests", SqlDbType.Bit).Value = !request.IsInternal;
                cmd.Parameters.Add("@SiteUrl", SqlDbType.NVarChar, 512).Value = NullableParam(Str(group, "SharePointSiteUrl"));
                await cmd.ExecuteNonQueryAsync(ct);
            }
        }
        catch
        {
            // Cache write is best-effort; the scheduled import will reconcile this row regardless.
        }
    }

    private static object NullableParam(string value) => value.Length > 0 ? value : DBNull.Value;

    // ----- logging -----

    private async Task<TeamCreationResult> FailAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, "FAIL", message, taskNumber, Severity.Error);
        return TeamCreationResult.Failed(correlationId, message);
    }

    private async Task LogAsync(Guid correlationId, string eventCode, string message, string taskNumber, Severity severity)
    {
        try
        {
            await _log.WriteAsync(new LogEntry
            {
                Area = "Teams",
                Action = "CreateTeam",
                EventCode = eventCode,
                Severity = severity,
                Message = message,
                TaskNumber = taskNumber,
                CorrelationId = correlationId,
                UserUpn = _auth.CurrentUser?.Upn,
            });
        }
        catch
        {
            // Logging is best-effort; never let it break the creation flow.
        }
    }

    private static void Progress(Action<string>? onProgress, string message) => onProgress?.Invoke(message);

    private static string Str(PSObject? o, string name)
        => o?.Properties[name]?.Value?.ToString() ?? "";
}
