using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Text.RegularExpressions;
using M365Manager.Core.M365;
using M365Manager.Core.PowerShell;
using M365Manager.Data.Logging;

namespace M365Manager.Core.Teams;

/// <summary>How the tenant decides which external domains Teams users may talk to.</summary>
public enum FederationMode
{
    /// <summary>Closed federation - only the domains on the allow list. Adding and removing is safe.</summary>
    AllowList,

    /// <summary>
    /// Open federation - every domain except the blocked ones. Adding a single domain to the allow
    /// list here switches the tenant to <see cref="AllowList"/> and cuts off everyone else.
    /// </summary>
    AllowAllKnownDomains,
}

public sealed class FederationState
{
    /// <summary>Master switch. False means no external federation at all, whatever the list says.</summary>
    public required bool AllowFederatedUsers { get; init; }
    public required FederationMode Mode { get; init; }

    /// <summary>Allowed partner domains, sorted. Empty in open mode.</summary>
    public required IReadOnlyList<string> AllowedDomains { get; init; }
}

public sealed class FederationChangeResult
{
    public required bool Succeeded { get; init; }

    /// <summary>Domains actually added or removed.</summary>
    public IReadOnlyList<string> Changed { get; init; } = Array.Empty<string>();

    /// <summary>Domains skipped because there was nothing to do (already on / not on the list).</summary>
    public IReadOnlyList<string> Skipped { get; init; } = Array.Empty<string>();

    public string? ErrorMessage { get; init; }

    public static FederationChangeResult Failed(string error) => new() { Succeeded = false, ErrorMessage = error };
}

/// <summary>
/// Teams external access (federation) with partner domains: reads the tenant federation
/// configuration and adds/removes domains on its allow list. There is no legacy script for this -
/// it used to be done in the Teams admin center - so every change goes to the audit log with its
/// ticket number like the rest of the app.
/// </summary>
public interface ITeamsFederationService
{
    Task<FederationState> GetStateAsync(CancellationToken ct = default);

    /// <summary>Adds partner domains to the allow list. Refuses in open mode (see <see cref="FederationMode"/>).</summary>
    Task<FederationChangeResult> AddDomainsAsync(string taskNumber, IReadOnlyList<string> domains, CancellationToken ct = default);

    Task<FederationChangeResult> RemoveDomainAsync(string taskNumber, string domain, CancellationToken ct = default);
}

public sealed partial class TeamsFederationService : ITeamsFederationService
{
    private const string Area = "TeamsFederation";

    private readonly PowerShellHost _host;
    private readonly TeamsPowerShellService _teams;
    private readonly IM365AuthService _auth;
    private readonly ILogService _log;

    public TeamsFederationService(PowerShellHost host, TeamsPowerShellService teams, IM365AuthService auth, ILogService log)
    {
        _host = host;
        _teams = teams;
        _auth = auth;
        _log = log;
    }

    /// <summary>
    /// AllowedDomains comes back as an object, not a list of strings: an AllowList whose
    /// AllowedDomain entries each carry a Domain property, or an AllowAllKnownDomains marker. Its
    /// text form ("Domain=a.com,Domain=b.com") is the fallback for module versions that serialize
    /// it differently. Flattened in the runspace so the app gets plain values back.
    /// </summary>
    private const string ReadScript = """
        $c = Get-CsTenantFederationConfiguration -ErrorAction Stop
        $ad = $c.AllowedDomains
        $text = "$ad"
        $list = @()
        if ($null -ne $ad -and $ad.PSObject.Properties['AllowedDomain']) {
            $list = @($ad.AllowedDomain | ForEach-Object { if ($_.PSObject.Properties['Domain']) { $_.Domain } else { "$_" } })
        }
        if ($list.Count -eq 0 -and $text -match 'Domain=') {
            $list = @([regex]::Matches($text, 'Domain=([^,\s]+)') | ForEach-Object { $_.Groups[1].Value })
        }
        [pscustomobject]@{
            AllowFederatedUsers = [bool]$c.AllowFederatedUsers
            AllowAll            = ($text -match 'AllowAllKnownDomains') -or ($null -ne $ad -and $ad.GetType().Name -match 'AllowAllKnownDomains')
            Domains             = @($list | Where-Object { $_ })
        }
        """;

    public async Task<FederationState> GetStateAsync(CancellationToken ct = default)
    {
        EnsureConnected();

        var result = await _host.InvokeAsync(ps => ps.AddScript(ReadScript), ct: ct);
        var o = result.FirstOrDefault()
                ?? throw new InvalidOperationException("Get-CsTenantFederationConfiguration returned nothing.");

        var domains = o.Properties["Domains"]?.Value is { } raw
            ? Enumerate(raw).Select(d => d.Trim().ToLowerInvariant()).Where(d => d.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(d => d, StringComparer.OrdinalIgnoreCase).ToList()
            : new List<string>();

        return new FederationState
        {
            AllowFederatedUsers = Bool(o, "AllowFederatedUsers"),
            Mode = Bool(o, "AllowAll") ? FederationMode.AllowAllKnownDomains : FederationMode.AllowList,
            AllowedDomains = domains,
        };
    }

    public async Task<FederationChangeResult> AddDomainsAsync(string taskNumber, IReadOnlyList<string> domains, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        try
        {
            if (string.IsNullOrWhiteSpace(taskNumber))
                return FederationChangeResult.Failed("Enter the ticket/task number authorizing this change.");
            if (domains.Count == 0)
                return FederationChangeResult.Failed("Enter at least one domain.");

            // Read fresh: the decision below depends on the mode the tenant is in right now, not
            // on whatever the page loaded earlier.
            var state = await GetStateAsync(ct);
            if (state.Mode == FederationMode.AllowAllKnownDomains)
            {
                return await FailAsync(correlationId, "AddFederatedDomain", taskNumber,
                    "The tenant allows all external domains (open federation). Adding a domain would switch it to an allow list " +
                    "and block every other external domain at once - nothing was changed. Change the mode in the Teams admin center if that is intended.");
            }

            var toAdd = domains.Where(d => !state.AllowedDomains.Contains(d, StringComparer.OrdinalIgnoreCase)).ToList();
            var skipped = domains.Except(toAdd, StringComparer.OrdinalIgnoreCase).ToList();

            await LogAsync(correlationId, "AddFederatedDomain", "INPUT",
                $"Add: {string.Join(", ", domains)}" + (skipped.Count > 0 ? $"; already allowed: {string.Join(", ", skipped)}" : ""),
                taskNumber, Severity.Info);

            if (toAdd.Count > 0)
            {
                // @{Add=...} changes only these entries; the rest of the list is left alone, so
                // nothing another admin added in the meantime is overwritten.
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-CsTenantFederationConfiguration")
                    .AddParameter("AllowedDomainsAsAList", new Hashtable { ["Add"] = new List<string>(toAdd) })
                    .AddParameter("ErrorAction", "Stop"), ct: ct);

                foreach (var domain in toAdd)
                    await LogAsync(correlationId, "AddFederatedDomain", "ADD", $"Allowed federation with {domain}", taskNumber, Severity.Success, domain);
            }

            return new FederationChangeResult { Succeeded = true, Changed = toAdd, Skipped = skipped };
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, "AddFederatedDomain", taskNumber, ex.Message);
        }
    }

    public async Task<FederationChangeResult> RemoveDomainAsync(string taskNumber, string domain, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        try
        {
            if (string.IsNullOrWhiteSpace(taskNumber))
                return FederationChangeResult.Failed("Enter the ticket/task number authorizing this change.");

            var state = await GetStateAsync(ct);
            if (!state.AllowedDomains.Contains(domain, StringComparer.OrdinalIgnoreCase))
                return new FederationChangeResult { Succeeded = true, Skipped = new[] { domain } };

            await LogAsync(correlationId, "RemoveFederatedDomain", "INPUT", $"Remove: {domain}", taskNumber, Severity.Info, domain);

            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-CsTenantFederationConfiguration")
                .AddParameter("AllowedDomainsAsAList", new Hashtable { ["Remove"] = new List<string> { domain } })
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            await LogAsync(correlationId, "RemoveFederatedDomain", "REMO", $"Removed federation with {domain}", taskNumber, Severity.Success, domain);
            return new FederationChangeResult { Succeeded = true, Changed = new[] { domain } };
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, "RemoveFederatedDomain", taskNumber, ex.Message, domain);
        }
    }

    /// <summary>
    /// Turns what the operator typed or pasted - "contoso.com", "https://www.contoso.com/",
    /// "user@contoso.com", one per line or comma-separated - into bare lowercase domains.
    /// Anything that still isn't a domain is reported rather than guessed at.
    /// </summary>
    public static (List<string> Domains, List<string> Invalid) ParseDomains(string input)
    {
        var domains = new List<string>();
        var invalid = new List<string>();

        foreach (var raw in (input ?? "").Split(new[] { ',', ';', '\n', '\r', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var value = raw.Trim().ToLowerInvariant();
            value = SchemePrefix().Replace(value, "");
            if (value.Contains('@'))
                value = value[(value.LastIndexOf('@') + 1)..];
            value = value.Split('/')[0].TrimEnd('.');
            if (value.StartsWith("www."))
                value = value[4..];

            // IDN (Umlaut domains) go to Teams in their ASCII form.
            try
            {
                value = new IdnMapping().GetAscii(value);
            }
            catch (ArgumentException)
            {
                invalid.Add(raw.Trim());
                continue;
            }

            if (DomainPattern().IsMatch(value))
            {
                if (!domains.Contains(value))
                    domains.Add(value);
            }
            else
            {
                invalid.Add(raw.Trim());
            }
        }

        return (domains, invalid);
    }

    [GeneratedRegex("^[a-z0-9]+://")]
    private static partial Regex SchemePrefix();

    /// <summary>At least one dot, labels of letters/digits/hyphens not starting or ending with a hyphen.</summary>
    [GeneratedRegex(@"^(?=.{4,253}$)([a-z0-9]([a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z0-9-]{2,63}$")]
    private static partial Regex DomainPattern();

    private void EnsureConnected()
    {
        if (!_teams.IsConnected)
            throw new InvalidOperationException(
                "Not connected to Microsoft Teams. Restart the app to sign in, or check the connection status on the dashboard.");
    }

    private static IEnumerable<string> Enumerate(object raw)
    {
        if (raw is PSObject wrapped)
            raw = wrapped.BaseObject;
        if (raw is string single)
            return new[] { single };
        if (raw is IEnumerable items)
            return items.Cast<object?>().Where(x => x is not null).Select(x => x!.ToString() ?? "");
        return new[] { raw.ToString() ?? "" };
    }

    private static bool Bool(PSObject o, string name) =>
        string.Equals(o.Properties[name]?.Value?.ToString(), "True", StringComparison.OrdinalIgnoreCase);

    private async Task<FederationChangeResult> FailAsync(Guid correlationId, string action, string taskNumber, string message, string? target = null)
    {
        await LogAsync(correlationId, action, "FAIL", message, taskNumber, Severity.Error, target);
        return FederationChangeResult.Failed(message);
    }

    private async Task LogAsync(Guid correlationId, string action, string eventCode, string message, string taskNumber, Severity severity,
        string? target = null)
    {
        try
        {
            await _log.WriteAsync(new LogEntry
            {
                Area = Area,
                Action = action,
                TargetObject = target,
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
            // Logging is best-effort; never let it break the change itself.
        }
    }
}
