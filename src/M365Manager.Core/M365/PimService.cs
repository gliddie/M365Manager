using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using M365Manager.Core.PowerShell;

namespace M365Manager.Core.M365;

/// <summary>One Entra role the signed-in admin is eligible for in PIM, and whether it is active now.</summary>
public sealed record PimRoleState(
    string RoleDefinitionId,
    string DisplayName,
    string DirectoryScopeId,
    bool IsActive,
    DateTimeOffset? ActiveUntil);

public interface IPimService
{
    /// <summary>The eligible roles found by the last check, active or not. Empty before the first one.</summary>
    IReadOnlyList<PimRoleState> Roles { get; }

    /// <summary>When the first activated role runs out; null if no time-limited role is active.</summary>
    DateTimeOffset? ActiveUntil { get; }

    /// <summary>Raised after every check or activation, on whatever thread did it.</summary>
    event EventHandler? StateChanged;
}

/// <summary>
/// Activates the signed-in admin's PIM roles - the step the team used to do by hand in the Entra
/// portal every morning, and again whenever the app had been left open overnight.
///
/// It is an <see cref="IM365Connector"/> between Graph (Order 0) and Exchange (2) on purpose:
/// Exchange Online fixes its cmdlet set at connect time from the roles active in that moment, so it
/// has to connect after the activation, never before.
///
/// Every role the admin is directly eligible for is activated - the app does not need to know which
/// roles exist, so a colleague with other roles than Global Administrator works without a change.
/// Roles granted through PIM-enabled groups are not covered (the team does not use them; that would
/// be a second API, /identityGovernance/privilegedAccess/group).
/// </summary>
public sealed class PimService : IPimService, IM365Connector
{
    private const string GraphBase = "https://graph.microsoft.com/v1.0";
    private const string GraphScope = "https://graph.microsoft.com/.default";

    /// <summary>
    /// Global Administrator needs no justification here, the other roles require one that nobody
    /// fills in meaningfully - so one fixed text satisfies every policy.
    /// </summary>
    private const string Justification = "Required for my daily work";

    /// <summary>Used when the role's policy cannot be read; the lowest maximum the team has (8 h).</summary>
    private const string FallbackDuration = "PT8H";

    private static readonly TimeSpan ActivationWait = TimeSpan.FromMinutes(3);
    private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(5);

    /// <summary>On a reconnect, roles with less than this left are deactivated and activated again.</summary>
    private static readonly TimeSpan RenewWindow = TimeSpan.FromMinutes(30);

    private static readonly HttpClient Http = new();

    private readonly IM365AuthService _auth;
    private readonly IPowerShellTranscript _transcript;
    private readonly Dictionary<string, string> _roleNames = new(StringComparer.OrdinalIgnoreCase);

    private bool _checked;
    private bool _renewRequested;

    public PimService(IM365AuthService auth, IPowerShellTranscript transcript)
    {
        _auth = auth;
        _transcript = transcript;
    }

    public IReadOnlyList<PimRoleState> Roles { get; private set; } = Array.Empty<PimRoleState>();

    public DateTimeOffset? ActiveUntil => Roles
        .Where(r => r.IsActive && r.ActiveUntil is not null)
        .Select(r => r.ActiveUntil)
        .Min();

    public event EventHandler? StateChanged;

    // ----- IM365Connector -----

    public string DisplayName => "Admin roles (PIM)";
    public int Order => 1;

    public bool IsConnected =>
        _checked && !_renewRequested
        && Roles.All(r => r.IsActive)
        && (ActiveUntil is null || ActiveUntil > DateTimeOffset.UtcNow.AddMinutes(2));

    public Task DisconnectAsync(CancellationToken ct = default)
    {
        _renewRequested = true;
        return Task.CompletedTask;
    }

    public async Task ConnectAsync(Action<string>? onPrompt = null, CancellationToken ct = default)
    {
        var me = _auth.CurrentUser
            ?? throw new InvalidOperationException("Sign in to Microsoft Graph first - PIM roles are activated for the signed-in admin.");

        await RefreshAsync(ct);

        var now = DateTimeOffset.UtcNow;
        var due = Roles
            .Where(r => !r.IsActive || (_renewRequested && r.ActiveUntil is { } until && until - now < RenewWindow))
            .ToList();

        if (due.Count > 0)
        {
            foreach (var role in due)
            {
                if (role.IsActive)
                {
                    // An activated role cannot be extended, only ended and activated again.
                    onPrompt?.Invoke($"Renewing {role.DisplayName}...");
                    await TryDeactivateAsync(me.Id, role, ct);
                }
                else
                {
                    onPrompt?.Invoke($"Activating {role.DisplayName}...");
                }

                await ActivateAsync(me.Id, role, ct);
            }

            await WaitUntilActiveAsync(due.Select(r => r.RoleDefinitionId).ToHashSet(StringComparer.OrdinalIgnoreCase), ct);

            // Tokens issued before the activation do not carry the new roles.
            _auth.InvalidateCachedTokens();
        }

        _renewRequested = false;
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    // ----- Reading the current state -----

    private async Task RefreshAsync(CancellationToken ct)
    {
        var eligible = await GetListAsync("/roleManagement/directory/roleEligibilityScheduleInstances/filterByCurrentUser(on='principal')", ct);
        var assigned = await GetListAsync("/roleManagement/directory/roleAssignmentScheduleInstances/filterByCurrentUser(on='principal')", ct);

        var roles = new List<PimRoleState>();
        foreach (var item in eligible)
        {
            var roleId = Text(item, "roleDefinitionId");
            if (string.IsNullOrEmpty(roleId))
                continue;
            var scope = Text(item, "directoryScopeId") is { Length: > 0 } s ? s : "/";

            // A permanent assignment of the same role counts as active too (no end date).
            var match = assigned.FirstOrDefault(a =>
                string.Equals(Text(a, "roleDefinitionId"), roleId, StringComparison.OrdinalIgnoreCase)
                && string.Equals(Text(a, "directoryScopeId") is { Length: > 0 } ds ? ds : "/", scope, StringComparison.OrdinalIgnoreCase));

            DateTimeOffset? until = null;
            if (match.ValueKind == JsonValueKind.Object
                && match.TryGetProperty("endDateTime", out var end) && end.ValueKind == JsonValueKind.String
                && DateTimeOffset.TryParse(end.GetString(), out var parsed))
                until = parsed;

            roles.Add(new PimRoleState(roleId, await GetRoleNameAsync(roleId, ct), scope,
                match.ValueKind == JsonValueKind.Object, until));
        }

        // The same role can be eligible twice (two schedules); one entry per role and scope is enough.
        Roles = roles
            .GroupBy(r => (r.RoleDefinitionId, r.DirectoryScopeId))
            .Select(g => g.OrderByDescending(r => r.IsActive).First())
            .ToList();
        _checked = true;
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    private async Task<string> GetRoleNameAsync(string roleId, CancellationToken ct)
    {
        if (_roleNames.TryGetValue(roleId, out var cached))
            return cached;

        var name = roleId;
        try
        {
            using var response = await SendAsync(HttpMethod.Get, $"/roleManagement/directory/roleDefinitions/{roleId}?$select=displayName", null, null, ct);
            if (response.Ok && Text(response.Json.RootElement, "displayName") is { Length: > 0 } display)
                name = display;
        }
        catch
        {
            // The id is a usable label if the name cannot be read.
        }

        _roleNames[roleId] = name;
        return name;
    }

    // ----- Activation -----

    private sealed record RolePolicy(string MaxDuration, string? AuthenticationContext);

    private async Task<RolePolicy> GetPolicyAsync(PimRoleState role, CancellationToken ct)
    {
        var maxDuration = FallbackDuration;
        string? authContext = null;
        try
        {
            var filter = Uri.EscapeDataString(
                $"scopeId eq '{role.DirectoryScopeId}' and scopeType eq 'DirectoryRole' and roleDefinitionId eq '{role.RoleDefinitionId}'");
            using var response = await SendAsync(HttpMethod.Get,
                $"/policies/roleManagementPolicyAssignments?$filter={filter}&$expand=policy($expand=rules)", null, null, ct);
            if (!response.Ok)
                throw new InvalidOperationException(response.Error);

            foreach (var assignment in Values(response.Json.RootElement))
            {
                if (!assignment.TryGetProperty("policy", out var policy) || !policy.TryGetProperty("rules", out var rules))
                    continue;

                foreach (var rule in rules.EnumerateArray())
                {
                    switch (Text(rule, "id"))
                    {
                        // "Activation maximum duration" from the role's PIM settings - 12 h for
                        // Global Administrator here, 8 h for the others.
                        case "Expiration_EndUser_Assignment" when Text(rule, "maximumDuration") is { Length: > 0 } duration:
                            maxDuration = duration;
                            break;
                        case "AuthenticationContext_EndUser_Assignment"
                            when rule.TryGetProperty("isEnabled", out var enabled) && enabled.ValueKind == JsonValueKind.True:
                            authContext = Text(rule, "claimValue");
                            break;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _transcript.Write(TranscriptLineKind.Error,
                $"Could not read the PIM policy of {role.DisplayName} ({ex.Message}) - activating for {FallbackDuration}.");
        }

        return new RolePolicy(maxDuration, authContext);
    }

    private async Task ActivateAsync(string principalId, PimRoleState role, CancellationToken ct)
    {
        var policy = await GetPolicyAsync(role, ct);
        var body = new
        {
            action = "selfActivate",
            principalId,
            roleDefinitionId = role.RoleDefinitionId,
            directoryScopeId = role.DirectoryScopeId,
            justification = Justification,
            scheduleInfo = new
            {
                startDateTime = DateTimeOffset.UtcNow.ToString("o"),
                expiration = new { type = "afterDuration", duration = policy.MaxDuration },
            },
        };

        const string path = "/roleManagement/directory/roleAssignmentScheduleRequests";
        var response = await SendAsync(HttpMethod.Post, path, body, null, ct);
        if (response.Ok)
            return;

        // The activation policy can demand a fresh MFA or an authentication context the sign-in
        // token does not carry yet. Graph names the failed rule; the matching claims request makes
        // Entra show the sign-in window with the step-up, and the request is sent once more.
        var claims = StepUpClaims(response, policy);
        if (claims is not null)
        {
            _transcript.Write(TranscriptLineKind.Output, $"{role.DisplayName}: the activation needs a fresh sign-in (MFA) - opening the sign-in window.");
            response = await SendAsync(HttpMethod.Post, path, body, claims, ct);
            if (response.Ok)
                return;
        }

        // Activated in the meantime (another window, the portal) - that is the goal anyway.
        if (response.Error.Contains("RoleAssignmentExists", StringComparison.OrdinalIgnoreCase))
            return;

        throw new InvalidOperationException($"Could not activate {role.DisplayName}: {response.Error}");
    }

    private static string? StepUpClaims(GraphResponse response, RolePolicy policy)
    {
        if (response.ClaimsChallenge is { Length: > 0 } challenge)
            return challenge;

        var error = response.Error;
        if ((error.Contains("AcrsRule", StringComparison.OrdinalIgnoreCase)
             || error.Contains("AcrsValidationFailed", StringComparison.OrdinalIgnoreCase))
            && policy.AuthenticationContext is { Length: > 0 } context)
            return "{\"access_token\":{\"acrs\":{\"essential\":true,\"value\":\"" + context + "\"}}}";

        if (error.Contains("MfaRule", StringComparison.OrdinalIgnoreCase))
            return "{\"access_token\":{\"amr\":{\"essential\":true,\"values\":[\"mfa\"]}}}";

        return null;
    }

    private async Task TryDeactivateAsync(string principalId, PimRoleState role, CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Post, "/roleManagement/directory/roleAssignmentScheduleRequests", new
        {
            action = "selfDeactivate",
            principalId,
            roleDefinitionId = role.RoleDefinitionId,
            directoryScopeId = role.DirectoryScopeId,
        }, null, ct);

        // PIM refuses to end an activation younger than five minutes. Then the activation below
        // fails with RoleAssignmentExists and the role simply keeps its current end time.
        if (!response.Ok)
        {
            _transcript.Write(TranscriptLineKind.Output, $"{role.DisplayName} could not be ended before renewing: {response.Error}");
            return;
        }

        // The deactivation is processed asynchronously too. Activating while the old activation is
        // still listed would be answered with RoleAssignmentExists - taken as success - and a moment
        // later the role would be gone. So wait until it has really ended.
        var deadline = DateTimeOffset.UtcNow + ActivationWait;
        while (DateTimeOffset.UtcNow < deadline)
        {
            await RefreshAsync(ct);
            if (!Roles.Any(r => r.RoleDefinitionId == role.RoleDefinitionId && r.DirectoryScopeId == role.DirectoryScopeId && r.IsActive))
                return;
            await Task.Delay(PollInterval, ct);
        }
    }

    /// <summary>
    /// The request returns before PIM has provisioned the role - polls until every requested role
    /// shows as active, so the connectors after this one sign in with the new rights.
    /// </summary>
    private async Task WaitUntilActiveAsync(HashSet<string> roleIds, CancellationToken ct)
    {
        var deadline = DateTimeOffset.UtcNow + ActivationWait;
        while (true)
        {
            await RefreshAsync(ct);
            var pending = Roles.Where(r => roleIds.Contains(r.RoleDefinitionId) && !r.IsActive).ToList();
            if (pending.Count == 0)
                return;

            if (DateTimeOffset.UtcNow >= deadline)
                throw new InvalidOperationException(
                    $"{string.Join(", ", pending.Select(r => r.DisplayName))} not active after {ActivationWait.TotalMinutes:0} minutes - " +
                    "PIM may still be processing. Try \"Renew roles & reconnect\" in a moment.");

            await Task.Delay(PollInterval, ct);
        }
    }

    // ----- Graph plumbing -----

    private sealed record GraphResponse(bool Ok, JsonDocument Json, string Error, string? ClaimsChallenge) : IDisposable
    {
        public void Dispose() => Json.Dispose();
    }

    private async Task<List<JsonElement>> GetListAsync(string path, CancellationToken ct)
    {
        using var response = await SendAsync(HttpMethod.Get, path, null, null, ct);
        if (!response.Ok)
            throw new InvalidOperationException($"Reading PIM roles failed: {response.Error}");
        return Values(response.Json.RootElement).Select(e => e.Clone()).ToList();
    }

    /// <summary>
    /// Its own request path rather than GraphRestClient: activation needs the status code, the
    /// claims challenge header and a token requested with claims, none of which that client exposes.
    /// </summary>
    private async Task<GraphResponse> SendAsync(HttpMethod method, string path, object? body, string? claims, CancellationToken ct)
    {
        var token = claims is null
            ? await _auth.GetAccessTokenAsync(GraphScope, ct)
            : await _auth.GetAccessTokenWithClaimsAsync(GraphScope, claims, ct);

        using var request = new HttpRequestMessage(method, GraphBase + path);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (body is not null)
            request.Content = JsonContent.Create(body);

        _transcript.Write(TranscriptLineKind.GraphRequest, $"{method.Method} {path}");

        using var response = await Http.SendAsync(request, ct);
        var payload = await response.Content.ReadAsStringAsync(ct);
        JsonDocument json;
        try { json = JsonDocument.Parse(string.IsNullOrWhiteSpace(payload) ? "{}" : payload); }
        catch (JsonException) { json = JsonDocument.Parse("{}"); }

        if (response.IsSuccessStatusCode)
            return new GraphResponse(true, json, "", null);

        var error = DescribeError(json, payload, (int)response.StatusCode);
        _transcript.Write(TranscriptLineKind.Error, $"Graph {method.Method} {path} failed: {error}");
        return new GraphResponse(false, json, error, ReadClaimsChallenge(response));
    }

    private static string DescribeError(JsonDocument json, string payload, int status)
    {
        if (json.RootElement.TryGetProperty("error", out var error))
        {
            var code = Text(error, "code");
            var message = Text(error, "message");
            return $"({status}) {code}: {message}".Trim();
        }
        return $"({status}) {(payload.Length > 300 ? payload[..300] + "..." : payload)}";
    }

    /// <summary>A WWW-Authenticate claims challenge, base64-decoded - present for authentication-context step-ups.</summary>
    private static string? ReadClaimsChallenge(HttpResponseMessage response)
    {
        foreach (var header in response.Headers.WwwAuthenticate)
        {
            var parameter = header.Parameter ?? "";
            var start = parameter.IndexOf("claims=\"", StringComparison.OrdinalIgnoreCase);
            if (start < 0)
                continue;
            start += "claims=\"".Length;
            var end = parameter.IndexOf('"', start);
            if (end < 0)
                continue;
            try
            {
                return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(parameter[start..end]));
            }
            catch (FormatException)
            {
                return null;
            }
        }
        return null;
    }

    private static IEnumerable<JsonElement> Values(JsonElement root) =>
        root.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array
            ? value.EnumerateArray()
            : Enumerable.Empty<JsonElement>();

    private static string Text(JsonElement element, string name) =>
        element.ValueKind == JsonValueKind.Object
        && element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : "";
}
