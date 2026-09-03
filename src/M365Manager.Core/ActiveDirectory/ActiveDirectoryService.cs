using System.DirectoryServices;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;

namespace M365Manager.Core.ActiveDirectory;

/// <summary>
/// See <see cref="IActiveDirectoryService"/>. Uses System.DirectoryServices with the process's own
/// credentials, so the bind is Kerberos as the signed-in admin - the legacy tool instead carried a
/// service account and a plaintext password in its source.
///
/// The search base is discovered from RootDSE rather than configured, so a machine that is joined to
/// the domain needs no setup at all. <see cref="NewHireSettings.LdapServer"/> only overrides which
/// controller is asked, for the case where the app runs somewhere without that discovery.
/// </summary>
public sealed class ActiveDirectoryService : IActiveDirectoryService
{
    private readonly ISettingsService _settings;
    private readonly IPowerShellTranscript _transcript;

    /// <summary>Signed and sealed so the attribute values never cross the wire in the clear.</summary>
    private const AuthenticationTypes BindAuth =
        AuthenticationTypes.Secure | AuthenticationTypes.Signing | AuthenticationTypes.Sealing;

    public ActiveDirectoryService(ISettingsService settings, IPowerShellTranscript transcript)
    {
        _settings = settings;
        _transcript = transcript;
    }

    public Task<AdUser?> FindUserAsync(string samAccountName, CancellationToken ct = default)
    {
        var sam = (samAccountName ?? "").Trim();
        if (sam.Length == 0)
            return Task.FromResult<AdUser?>(null);

        // System.DirectoryServices is synchronous and blocking; keep it off the UI thread.
        return Task.Run(() =>
        {
            _transcript.Write(TranscriptLineKind.Command, $"[LDAP] search (&(objectCategory=person)(objectClass=user)(sAMAccountName={sam}))");

            using var root = CreateSearchRoot();
            using var searcher = new DirectorySearcher(root)
            {
                Filter = $"(&(objectCategory=person)(objectClass=user)(sAMAccountName={Escape(sam)}))",
                SearchScope = SearchScope.Subtree,
                SizeLimit = 2,
                ClientTimeout = TimeSpan.FromSeconds(30),
            };

            searcher.PropertiesToLoad.AddRange(new[]
            {
                "distinguishedName", "sAMAccountName", "displayName", "userPrincipalName",
                "mail", "physicalDeliveryOfficeName", "givenName",
                "msRTCSIP-Line", "msRTCSIP-PrimaryUserAddress",
            });

            using var results = searcher.FindAll();
            var found = results.Cast<SearchResult>().FirstOrDefault();
            if (found is null)
            {
                _transcript.Write(TranscriptLineKind.Output, $"[LDAP] no account found for '{sam}'");
                return null;
            }

            var user = new AdUser(
                DistinguishedName: Value(found, "distinguishedName"),
                SamAccountName: Value(found, "sAMAccountName"),
                DisplayName: Value(found, "displayName"),
                UserPrincipalName: Value(found, "userPrincipalName"),
                Mail: Value(found, "mail"),
                Office: Value(found, "physicalDeliveryOfficeName"),
                GivenName: Value(found, "givenName"),
                LineUri: NullIfEmpty(Value(found, "msRTCSIP-Line")),
                SipAddress: NullIfEmpty(Value(found, "msRTCSIP-PrimaryUserAddress")));

            _transcript.Write(TranscriptLineKind.Output,
                $"[LDAP] {user.DisplayName} <{user.UserPrincipalName}> office='{user.Office}' line='{user.LineUri ?? "(none)"}'");

            return user;
        }, ct);
    }

    public Task<IReadOnlyList<AdAttributeResult>> WriteTelephonyAttributesAsync(
        AdUser user,
        string lineUri,
        string sipAddress,
        string deploymentLocator,
        CancellationToken ct = default)
    {
        // Order matters only in that the two identity attributes - the ones Entra Connect turns into
        // the cloud number and SIP address - go first. If the run dies after those, the account is
        // still consistent with what Teams will be told.
        var writes = new (string Attribute, string Value)[]
        {
            ("msRTCSIP-Line", lineUri),
            ("msRTCSIP-PrimaryUserAddress", sipAddress),
            ("msRTCSIP-DeploymentLocator", deploymentLocator),
            ("msRTCSIP-FederationEnabled", "TRUE"),
            ("msRTCSIP-InternetAccessEnabled", "TRUE"),
            ("msRTCSIP-UserEnabled", "TRUE"),
        };

        return Task.Run<IReadOnlyList<AdAttributeResult>>(() =>
        {
            var results = new List<AdAttributeResult>(writes.Length);

            using var entry = CreateEntry($"LDAP://{ServerPrefix()}{user.DistinguishedName}");

            foreach (var (attribute, value) in writes)
            {
                ct.ThrowIfCancellationRequested();
                _transcript.Write(TranscriptLineKind.Command, $"[LDAP] set {attribute} = {value} on {user.SamAccountName}");

                try
                {
                    // Committing per attribute rather than once at the end: a value the schema or
                    // ACL rejects then fails alone instead of discarding the whole batch, which is
                    // what makes the per-attribute result below meaningful.
                    entry.Properties[attribute].Value = value;
                    entry.CommitChanges();
                    results.Add(new AdAttributeResult(attribute, value, null));
                }
                catch (Exception ex)
                {
                    // Drop the rejected value so the next CommitChanges doesn't retry it and fail again.
                    entry.RefreshCache();
                    _transcript.Write(TranscriptLineKind.Error, $"[LDAP] {attribute} failed: {ex.Message}");
                    results.Add(new AdAttributeResult(attribute, value, ex.Message));
                }
            }

            return results;
        }, ct);
    }

    /// <summary>
    /// The domain's default naming context, so no search base has to be configured. Falls back to
    /// the configured server when one is set.
    /// </summary>
    private DirectoryEntry CreateSearchRoot()
    {
        using var rootDse = CreateEntry($"LDAP://{ServerPrefix()}RootDSE");

        var namingContext = rootDse.Properties["defaultNamingContext"].Value as string;
        if (string.IsNullOrWhiteSpace(namingContext))
            throw new InvalidOperationException(
                "Could not read the domain naming context from Active Directory. If this machine is not domain-joined, "
                + "set a domain controller under Settings > New Hire.");

        return CreateEntry($"LDAP://{ServerPrefix()}{namingContext}");
    }

    private DirectoryEntry CreateEntry(string path) =>
        // No username/password: binds as the process identity, i.e. the signed-in admin.
        new(path, null, null, BindAuth);

    /// <summary>"dc01.corp.contoso.com/" when a controller is configured, empty for serverless bind.</summary>
    private string ServerPrefix()
    {
        var server = _settings.Current.NewHire.LdapServer?.Trim() ?? "";
        return server.Length == 0 ? "" : $"{server.TrimEnd('/')}/";
    }

    private static string Value(SearchResult result, string property) =>
        result.Properties.Contains(property) && result.Properties[property].Count > 0
            ? result.Properties[property][0]?.ToString() ?? ""
            : "";

    private static string? NullIfEmpty(string value) => value.Length == 0 ? null : value;

    /// <summary>
    /// Escapes the LDAP filter metacharacters from RFC 4515, so a SamAccountName carrying one of
    /// them can't rewrite the filter.
    /// </summary>
    private static string Escape(string value)
    {
        var escaped = new System.Text.StringBuilder(value.Length);
        foreach (var c in value)
        {
            escaped.Append(c switch
            {
                '\\' => "\\5c",
                '*' => "\\2a",
                '(' => "\\28",
                ')' => "\\29",
                '\0' => "\\00",
                '/' => "\\2f",
                _ => c.ToString(),
            });
        }

        return escaped.ToString();
    }
}
