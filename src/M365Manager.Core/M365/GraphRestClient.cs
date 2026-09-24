using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using M365Manager.Core.PowerShell;

namespace M365Manager.Core.M365;

/// <summary>
/// Thin wrapper over the Graph v1.0 REST API for endpoints the Kiota SDK models awkwardly.
///
/// TeamsService and SharedMailboxService each still carry their own private copy of this - they
/// predate it and are in active use, so they were left alone; new callers should use this instead.
/// </summary>
public sealed class GraphRestClient
{
    private const string BaseUrl = "https://graph.microsoft.com/v1.0";
    private const string Scope = "https://graph.microsoft.com/.default";
    private static readonly HttpClient Http = new();

    private readonly IM365AuthService _auth;
    private readonly IPowerShellTranscript _transcript;

    public GraphRestClient(IM365AuthService auth, IPowerShellTranscript transcript)
    {
        _auth = auth;
        _transcript = transcript;
    }

    public Task<JsonDocument> GetAsync(string path, bool advancedQuery = false, CancellationToken ct = default)
        => SendAsync(HttpMethod.Get, path, null, advancedQuery, ct);

    public Task<JsonDocument> PostAsync(string path, object body, CancellationToken ct = default)
        => SendAsync(HttpMethod.Post, path, body, false, ct);

    public Task<JsonDocument> PatchAsync(string path, object body, CancellationToken ct = default)
        => SendAsync(HttpMethod.Patch, path, body, false, ct);

    /// <summary>PUT, e.g. teamifying a group via /groups/{id}/team.</summary>
    public Task<JsonDocument> PutAsync(string path, object body, CancellationToken ct = default)
        => SendAsync(HttpMethod.Put, path, body, false, ct);

    /// <summary>DELETE, e.g. removing a group member via /groups/{id}/members/{userId}/$ref.</summary>
    public Task<JsonDocument> DeleteAsync(string path, CancellationToken ct = default)
        => SendAsync(HttpMethod.Delete, path, null, false, ct);

    /// <summary>
    /// <paramref name="advancedQuery"/> adds ConsistencyLevel: eventual, which Graph requires for
    /// $filter on properties outside the default filterable set (companyName among them).
    /// </summary>
    private async Task<JsonDocument> SendAsync(HttpMethod method, string path, object? body, bool advancedQuery, CancellationToken ct)
    {
        var token = await _auth.GetAccessTokenAsync(Scope, ct);

        using var request = new HttpRequestMessage(method, path.StartsWith("http") ? path : BaseUrl + path);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        if (advancedQuery)
            request.Headers.Add("ConsistencyLevel", "eventual");
        if (body is not null)
            request.Content = JsonContent.Create(body);

        // Echoed to the same console the PowerShell side writes to. The bearer token is never part
        // of this - only the method, the path and the request body.
        WriteToTranscript(method, path, body);

        using var response = await Http.SendAsync(request, ct);
        var payload = await response.Content.ReadAsStringAsync(ct);

        if (!response.IsSuccessStatusCode)
        {
            var summary = Summarize(payload);
            _transcript.Write(TranscriptLineKind.Error, $"Graph {method} {path} failed ({(int)response.StatusCode}): {summary}");
            throw new InvalidOperationException($"Graph {method} {path} failed ({(int)response.StatusCode}): {summary}");
        }

        // 204 No Content (a successful PATCH) has nothing to parse.
        return string.IsNullOrWhiteSpace(payload)
            ? JsonDocument.Parse("{}")
            : JsonDocument.Parse(payload);
    }

    /// <summary>
    /// Renders the request into the console the way the PowerShell side renders a cmdlet. Bodies are
    /// included because that is what makes a POST readable - and capped, because a sendMail body can
    /// carry a whole embedded image.
    /// </summary>
    private void WriteToTranscript(HttpMethod method, string path, object? body)
    {
        _transcript.Write(TranscriptLineKind.GraphRequest, $"{method.Method} {path}");

        if (body is null)
            return;

        try
        {
            var json = JsonSerializer.Serialize(body, new JsonSerializerOptions { WriteIndented = true });
            _transcript.Write(TranscriptLineKind.Output, json.Length > 2000 ? json[..2000] + "\n... (truncated)" : json);
        }
        catch
        {
            // A body that will not serialize for display must not stop the call itself.
        }
    }

    /// <summary>Pulls Graph's nested error message out of the response so the UI shows the cause, not the envelope.</summary>
    private static string Summarize(string payload)
    {
        if (string.IsNullOrWhiteSpace(payload))
            return "(no response body)";

        try
        {
            using var json = JsonDocument.Parse(payload);
            if (json.RootElement.TryGetProperty("error", out var error)
                && error.TryGetProperty("message", out var message))
            {
                var text = message.GetString() ?? payload;
                return error.TryGetProperty("code", out var code)
                    ? $"{code.GetString()}: {text}"
                    : text;
            }
        }
        catch (JsonException)
        {
            // not JSON - fall through to the raw text
        }

        return payload.Length > 500 ? payload[..500] + "..." : payload;
    }

    /// <summary>Follows @odata.nextLink, returning every "value" element across all pages.</summary>
    public async Task<List<JsonElement>> GetAllPagesAsync(string path, bool advancedQuery = false, int maxPages = 50, CancellationToken ct = default)
    {
        var results = new List<JsonElement>();
        var next = path;

        for (var page = 0; page < maxPages && next is not null; page++)
        {
            using var json = await GetAsync(next, advancedQuery, ct);

            if (json.RootElement.TryGetProperty("value", out var value) && value.ValueKind == JsonValueKind.Array)
            {
                // The JsonDocument is disposed at the end of this iteration, so each element has to
                // be cloned out of it before it goes away.
                foreach (var item in value.EnumerateArray())
                    results.Add(item.Clone());
            }

            next = json.RootElement.TryGetProperty("@odata.nextLink", out var link) ? link.GetString() : null;
        }

        return results;
    }
}
