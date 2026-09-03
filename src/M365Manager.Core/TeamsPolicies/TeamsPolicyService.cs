using System.Text.Json;
using M365Manager.Core.M365;
using M365Manager.Core.Settings;
using M365Manager.Data.Logging;

namespace M365Manager.Core.TeamsPolicies;

/// <summary>
/// See <see cref="ITeamsPolicyService"/>. Pure Graph, like TraceGuestService - no Exchange
/// PowerShell, so it works whether or not the Exchange connection came up.
///
/// No new Graph consent is needed: reading groups and a user's memberships is covered by
/// Group.ReadWrite.All / User.Read.All, and changing membership by GroupMember.ReadWrite.All - all
/// already requested (see M365AuthService.Scopes).
/// </summary>
public sealed class TeamsPolicyService : ITeamsPolicyService
{
    private readonly GraphRestClient _graph;
    private readonly ISettingsService _settings;
    private readonly IM365AuthService _auth;
    private readonly ILogService _log;

    public TeamsPolicyService(GraphRestClient graph, ISettingsService settings, IM365AuthService auth, ILogService log)
    {
        _graph = graph;
        _settings = settings;
        _auth = auth;
        _log = log;
    }

    // ----- reading -----

    public async Task<IReadOnlyList<PolicyCategory>> GetCategoriesAsync(CancellationToken ct = default)
    {
        var categories = new List<PolicyCategory>();

        foreach (var prefix in _settings.Current.TeamsPolicies.Prefixes)
        {
            categories.Add(new PolicyCategory
            {
                Prefix = prefix,
                AvailableGroups = await GetGroupsWithPrefixAsync(prefix, ct),
            });
        }

        return categories;
    }

    public async Task<UserPolicyLookup> LookupUserAsync(string identity, CancellationToken ct = default)
    {
        var value = (identity ?? "").Trim();
        if (value.Length == 0)
            return UserPolicyLookup.Failed("Enter a user to look up.");

        var prefixes = _settings.Current.TeamsPolicies.Prefixes;
        if (prefixes.Count == 0)
            return UserPolicyLookup.Failed("No policy group prefixes configured - set them under Settings > Teams policies.");

        PolicyUser user;
        try
        {
            user = await ResolveUserAsync(value, ct);
        }
        catch (Exception ex)
        {
            return UserPolicyLookup.Failed($"'{value}' could not be resolved to a user: {ex.Message}");
        }

        // Direct memberships only. Teams evaluates group policy assignment on the groups a user is
        // actually a member of, and only a direct membership is something this page can remove.
        var memberOf = await _graph.GetAllPagesAsync(
            $"/users/{Uri.EscapeDataString(user.Id)}/memberOf?$select=id,displayName,description", ct: ct);

        var categories = new List<PolicyCategory>();
        foreach (var prefix in prefixes)
        {
            var current = memberOf
                .Select(g => ToGroup(g, prefix))
                .Where(g => g is not null && g.DisplayName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                .Select(g => g!)
                .OrderBy(g => g.DisplayName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            categories.Add(new PolicyCategory
            {
                Prefix = prefix,
                AvailableGroups = await GetGroupsWithPrefixAsync(prefix, ct),
                CurrentGroups = current,
            });
        }

        return new UserPolicyLookup { Succeeded = true, User = user, Categories = categories };
    }

    private async Task<IReadOnlyList<TeamsPolicyGroup>> GetGroupsWithPrefixAsync(string prefix, CancellationToken ct)
    {
        // Single quotes have to be doubled inside an OData string literal.
        var literal = prefix.Replace("'", "''");
        var path = $"/groups?$filter=startswith(displayName,'{Uri.EscapeDataString(literal)}')"
                   + "&$select=id,displayName,description&$top=999";

        var groups = await _graph.GetAllPagesAsync(path, ct: ct);

        return groups
            .Select(g => ToGroup(g, prefix))
            .Where(g => g is not null)
            .Select(g => g!)
            .OrderBy(g => g.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static TeamsPolicyGroup? ToGroup(JsonElement element, string prefix)
    {
        var id = Text(element, "id");
        var name = Text(element, "displayName");
        return id.Length > 0 && name.Length > 0
            ? new TeamsPolicyGroup(id, name, Text(element, "description"), prefix)
            : null;
    }

    private async Task<PolicyUser> ResolveUserAsync(string identity, CancellationToken ct)
    {
        var candidate = AppendDomain(identity);

        using var json = await _graph.GetAsync(
            $"/users/{Uri.EscapeDataString(candidate)}?$select=id,displayName,userPrincipalName,mail", ct: ct);

        var root = json.RootElement;
        return new PolicyUser(
            Text(root, "id"),
            Text(root, "displayName"),
            Text(root, "userPrincipalName"),
            Text(root, "mail"));
    }

    /// <summary>A bare SamAccountName is completed with the configured sign-in domain, as elsewhere.</summary>
    private string AppendDomain(string identity)
    {
        var value = identity.Trim();
        if (value.Contains('@'))
            return value;

        var domain = _settings.Current.M365.Domain.Trim();
        return domain.Length == 0 ? value : $"{value}@{domain.TrimStart('@')}";
    }

    // ----- changing -----

    public async Task<AssignPolicyResult> AssignAsync(AssignPolicyRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "STAR", "Teams policy assignment has started", request, Severity.Info);

            if (string.IsNullOrWhiteSpace(request.TaskNumber))
                return await FailAsync(correlationId, request, "Enter the ticket/task number authorizing this change.");

            // Anything already correct is left alone, so re-applying the same group is a no-op
            // rather than a remove-then-add that briefly drops the user's policy.
            var toRemove = request.Replacing
                .Where(g => request.Group is null || !string.Equals(g.Id, request.Group.Id, StringComparison.OrdinalIgnoreCase))
                .ToList();

            var alreadyAssigned = request.Group is not null
                && request.Replacing.Any(g => string.Equals(g.Id, request.Group.Id, StringComparison.OrdinalIgnoreCase));

            if (request.Group is null && toRemove.Count == 0)
                return await FailAsync(correlationId, request, "Nothing to do - the user holds no group under this prefix.");

            await LogAsync(correlationId, "INPUT",
                $"User={request.User.Upn}; Prefix={request.Prefix}; "
                + $"Assign={(request.Group?.DisplayName ?? "(none)")}; "
                + $"Removing={(toRemove.Count == 0 ? "(none)" : string.Join(", ", toRemove.Select(g => g.DisplayName)))}",
                request, Severity.Info);

            var added = new List<string>();
            var removed = new List<string>();
            var warnings = new List<string>();

            // Add first: if the removal then fails the user is over-assigned, which Teams resolves
            // by policy precedence. The other order could leave them with no policy at all.
            if (request.Group is not null && !alreadyAssigned)
            {
                Progress(onProgress, $"Adding {request.User.DisplayName} to {request.Group.DisplayName}...");
                var body = new Dictionary<string, object?>
                {
                    ["@odata.id"] = $"https://graph.microsoft.com/v1.0/directoryObjects/{request.User.Id}",
                };
                using var _ = await _graph.PostAsync($"/groups/{request.Group.Id}/members/$ref", body, ct);

                added.Add(request.Group.DisplayName);
                await LogAsync(correlationId, "ADD", $"Added {request.User.Upn} to {request.Group.DisplayName}", request, Severity.Success);
            }
            else if (alreadyAssigned)
            {
                await LogAsync(correlationId, "SKIP", $"{request.User.Upn} is already a member of {request.Group!.DisplayName}", request, Severity.Info);
            }

            foreach (var group in toRemove)
            {
                Progress(onProgress, $"Removing {request.User.DisplayName} from {group.DisplayName}...");
                try
                {
                    using var _ = await _graph.DeleteAsync($"/groups/{group.Id}/members/{request.User.Id}/$ref", ct);
                    removed.Add(group.DisplayName);
                    await LogAsync(correlationId, "REMO", $"Removed {request.User.Upn} from {group.DisplayName}", request, Severity.Success);
                }
                catch (Exception ex)
                {
                    warnings.Add($"{group.DisplayName} could not be removed: {ex.Message}");
                    await LogAsync(correlationId, "WARN", $"Could not remove {request.User.Upn} from {group.DisplayName}: {ex.Message}", request, Severity.Warning);
                }
            }

            await LogAsync(correlationId, "DONE", "Teams policy assignment has completed", request, Severity.Success);

            return new AssignPolicyResult
            {
                Succeeded = true,
                Added = added,
                Removed = removed,
                CorrelationId = correlationId,
                WarningMessage = warnings.Count == 0
                    ? null
                    : $"The user now holds more than one group under '{request.Prefix}': {string.Join(" ", warnings)}",
            };
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, request, ex.Message);
        }
    }

    // ----- helpers -----

    private static string Text(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : "";

    private static void Progress(Action<string>? onProgress, string message) => onProgress?.Invoke(message);

    private async Task<AssignPolicyResult> FailAsync(Guid correlationId, AssignPolicyRequest request, string error)
    {
        await LogAsync(correlationId, "FAIL", error, request, Severity.Error);
        return AssignPolicyResult.Failed(correlationId, error);
    }

    private async Task LogAsync(Guid correlationId, string eventCode, string message, AssignPolicyRequest request, Severity severity)
    {
        try
        {
            await _log.WriteAsync(new LogEntry
            {
                CorrelationId = correlationId,
                Area = "TeamsPolicies",
                Action = "AssignPolicy",
                EventCode = eventCode,
                Severity = severity,
                TargetObject = request.User.Upn,
                TaskNumber = request.TaskNumber,
                UserUpn = _auth.CurrentUser?.Upn,
                Message = message,
            });
        }
        catch
        {
            // Logging is best-effort - SQL may not be reachable.
        }
    }
}
