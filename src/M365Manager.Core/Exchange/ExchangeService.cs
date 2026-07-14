using System.Collections;
using System.Data;
using System.Management.Automation;
using M365Manager.Core.M365;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Data.Logging;
using Microsoft.Data.SqlClient;

namespace M365Manager.Core.Exchange;

/// <summary>
/// Exchange Online operations via the hosted PowerShell runspace (bundled EXO module).
/// Connects with an access token from the signed-in admin, so actions run under their identity.
/// </summary>
public sealed class ExchangeService : IExchangeService, IM365Connector
{
    private readonly PowerShellHost _host;
    private readonly IM365AuthService _auth;
    private readonly ISettingsService _settings;

    public ExchangeService(PowerShellHost host, IM365AuthService auth, ISettingsService settings)
    {
        _host = host;
        _auth = auth;
        _settings = settings;
    }

    public string DisplayName => "Exchange Online";
    public int Order => 1;

    public bool IsConnected { get; private set; }

    public async Task ConnectAsync(Action<string>? onPrompt = null, CancellationToken ct = default)
    {
        await _host.EnsureExchangeModuleAsync(ct);

        var upn = _auth.CurrentUser?.Upn;

        // Device-code auth: delegated (keeps the colleague's identity) and works in a
        // hosted GUI process without a WAM window handle.
        await _host.InvokeAsync(ps =>
        {
            ps.AddCommand("Connect-ExchangeOnline")
              .AddParameter("Device", true)
              .AddParameter("ShowBanner", false)
              .AddParameter("ErrorAction", "Stop");
            if (!string.IsNullOrWhiteSpace(upn))
                ps.AddParameter("UserPrincipalName", upn);
        }, onPrompt, ct);

        IsConnected = true;
    }

    public async Task<IReadOnlyList<DistributionGroupInfo>> SearchGroupsAsync(string search, CancellationToken ct = default)
    {
        var cached = await SearchGroupsCachedAsync(search, ct);
        if (cached.Count > 0)
            return cached;

        var term = (search ?? "").Replace("'", "").Replace("\"", "").Trim();
        var pattern = $"*{term}*";

        // Distribution + mail-enabled security groups
        var dg = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-DistributionGroup")
            .AddParameter("Identity", pattern)
            .AddParameter("ResultSize", 50)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

        // Microsoft 365 (unified) groups
        var ug = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-UnifiedGroup")
            .AddParameter("Identity", pattern)
            .AddParameter("ResultSize", 50)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

        return dg.Concat(ug)
            .Select(MapGroup)
            .GroupBy(g => g.PrimarySmtpAddress)
            .Select(grp => grp.First())
            .OrderBy(g => g.DisplayName)
            .ToList();
    }

    public async Task<IReadOnlyList<DistributionGroupInfo>> SearchGroupsCachedAsync(string search, CancellationToken ct = default)
    {
        var sql = _settings.Current.Sql;
        if (string.IsNullOrWhiteSpace(sql.Server) || string.IsNullOrWhiteSpace(sql.Database))
            return Array.Empty<DistributionGroupInfo>();

        var term = (search ?? "").Trim();
        if (string.IsNullOrWhiteSpace(term))
            return Array.Empty<DistributionGroupInfo>();

        await using var conn = new SqlConnection(BuildConnectionString(sql));
        await conn.OpenAsync(ct);

        var query = @"
SELECT TOP 50
    Id,
    DisplayName,
    PrimarySmtpAddress,
    Alias,
    GroupTypes,
    RecipientTypeDetails,
    Notes,
    ManagedBy,
    IsDeletedInM365,
    LastSeenAtUtc,
    MailEnabled,
    SecurityEnabled,
    IsMembershipDynamic
FROM dbo.Groups
WHERE LOWER(DisplayName) LIKE @term
   OR LOWER(PrimarySmtpAddress) LIKE @term
   OR LOWER(Alias) LIKE @term
   OR LOWER(MailNickname) LIKE @term
ORDER BY DisplayName;
";

        await using var cmd = new SqlCommand(query, conn);
        cmd.Parameters.Add("@term", SqlDbType.NVarChar, 4000).Value = $"%{term.ToLowerInvariant()}%";

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        var list = new List<DistributionGroupInfo>();
        while (await reader.ReadAsync(ct))
        {
            list.Add(new DistributionGroupInfo
            {
                Name = reader.IsDBNull(0) ? "" : reader.GetGuid(0).ToString(),
                DisplayName = reader.IsDBNull(1) ? "" : reader.GetString(1),
                PrimarySmtpAddress = reader.IsDBNull(2) ? "" : reader.GetString(2),
                Alias = reader.IsDBNull(3) ? "" : reader.GetString(3),
                GroupType = reader.IsDBNull(4) ? "" : reader.GetString(4),
                RecipientTypeDetails = reader.IsDBNull(5) ? "" : reader.GetString(5),
                Notes = reader.IsDBNull(6) ? "" : reader.GetString(6),
                ManagedBy = reader.IsDBNull(7) ? new List<string>() : reader.GetString(7).Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList(),
                IsDeletedInM365 = !reader.IsDBNull(8) && reader.GetBoolean(8),
                LastSeenAtUtc = reader.IsDBNull(9) ? null : reader.GetDateTime(9),
                MailEnabled = !reader.IsDBNull(10) && reader.GetBoolean(10),
                SecurityEnabled = !reader.IsDBNull(11) && reader.GetBoolean(11),
                IsMembershipDynamic = !reader.IsDBNull(12) && reader.GetBoolean(12),
            });
        }

        return list;
    }

    public async Task<IReadOnlyList<GroupMemberInfo>> GetMembersAsync(DistributionGroupInfo group, CancellationToken ct = default)
    {
        var sql = _settings.Current.Sql;
        if (!string.IsNullOrWhiteSpace(sql.Server) && !string.IsNullOrWhiteSpace(sql.Database))
        {
            var cached = await GetMembersCachedAsync(group, ct);
            if (cached.Count > 0)
                return cached;
        }

        var identity = string.IsNullOrWhiteSpace(group.PrimarySmtpAddress) ? group.Name : group.PrimarySmtpAddress;
        var results = await GetLiveMembersAsync(group, identity, ct);
        return MapMembers(results);
    }

    /// <summary>
    /// Re-fetches direct members live and overwrites the cached direct-member rows (IsNested = 0)
    /// for this group, so an add/remove made here or in another tool shows up without waiting for
    /// the next scheduled import. Nested-member rows (produced only by the import script's
    /// recursive resolution) are left untouched since a direct-member fetch can't see them.
    /// </summary>
    public async Task<IReadOnlyList<GroupMemberInfo>> SyncMembersAsync(DistributionGroupInfo group, CancellationToken ct = default)
    {
        var identity = string.IsNullOrWhiteSpace(group.PrimarySmtpAddress) ? group.Name : group.PrimarySmtpAddress;
        var results = await GetLiveMembersAsync(group, identity, ct);

        await TryWriteDirectMembersToCacheAsync(group, results, ct);

        return MapMembers(results);
    }

    private async Task<IReadOnlyList<PSObject>> GetLiveMembersAsync(DistributionGroupInfo group, string identity, CancellationToken ct)
    {
        var results = group.IsUnifiedGroup
            ? await _host.InvokeAsync(ps => ps
                .AddCommand("Get-UnifiedGroupLinks")
                .AddParameter("Identity", identity)
                .AddParameter("LinkType", "Members")
                .AddParameter("ResultSize", "Unlimited")
                .AddParameter("ErrorAction", "Stop"), ct: ct)
            : await _host.InvokeAsync(ps => ps
                .AddCommand("Get-DistributionGroupMember")
                .AddParameter("Identity", identity)
                .AddParameter("ResultSize", "Unlimited")
                .AddParameter("ErrorAction", "Stop"), ct: ct);

        await EnrichWithUserPrincipalNamesAsync(results, ct);
        return results;
    }

    /// <summary>
    /// Get-DistributionGroupMember/Get-UnifiedGroupLinks return recipient stubs with no
    /// UserPrincipalName property at all - that's only present on Get-User/Get-Mailbox objects.
    /// Resolve it with one batched Get-User lookup (by directory id, falling back to SMTP
    /// address) and stamp it onto each result as a note property so downstream mapping/caching
    /// can read it like any other field.
    /// </summary>
    private async Task EnrichWithUserPrincipalNamesAsync(IReadOnlyList<PSObject> members, CancellationToken ct)
    {
        var lookupKeys = members
            .Select(MemberLookupKey)
            .Where(k => k.Length > 0)
            .Distinct()
            .ToArray();

        if (lookupKeys.Length == 0)
            return;

        const string script = @"
param($Ids)
foreach ($id in $Ids) {
    try {
        $u = Get-User -Identity $id -ErrorAction Stop
        [pscustomobject]@{ Id = $id; Upn = $u.UserPrincipalName }
    } catch { }
}
";

        var resolved = await _host.InvokeAsync(ps => ps
            .AddScript(script)
            .AddParameter("Ids", lookupKeys), ct: ct);

        var upnByKey = resolved
            .Select(o => (Id: Str(o, "Id"), Upn: Str(o, "Upn")))
            .Where(x => x.Id.Length > 0 && x.Upn.Length > 0)
            .ToDictionary(x => x.Id, x => x.Upn, StringComparer.OrdinalIgnoreCase);

        foreach (var member in members)
        {
            var key = MemberLookupKey(member);
            if (key.Length > 0 && upnByKey.TryGetValue(key, out var upn) && member.Properties["UserPrincipalName"] is null)
                member.Properties.Add(new PSNoteProperty("UserPrincipalName", upn));
        }
    }

    private static string MemberLookupKey(PSObject o)
    {
        var id = Str(o, "ExternalDirectoryObjectId");
        return id.Length > 0 ? id : Str(o, "PrimarySmtpAddress");
    }

    private static List<GroupMemberInfo> MapMembers(IReadOnlyList<PSObject> results) => results.Select(o => new GroupMemberInfo
    {
        DisplayName = Str(o, "DisplayName"),
        PrimarySmtpAddress = Str(o, "PrimarySmtpAddress"),
        UserPrincipalName = Str(o, "UserPrincipalName"),
        RecipientType = Str(o, "RecipientTypeDetails"),
    })
    .OrderBy(m => m.DisplayName)
    .ToList();

    private async Task TryWriteDirectMembersToCacheAsync(DistributionGroupInfo group, IReadOnlyList<PSObject> liveMembers, CancellationToken ct)
    {
        var sql = _settings.Current.Sql;
        if (string.IsNullOrWhiteSpace(sql.Server) || string.IsNullOrWhiteSpace(sql.Database))
            return;
        if (!Guid.TryParse(group.Name, out var groupId))
            return; // Group wasn't sourced from the cache (no known Groups.Id) - nothing to sync into.

        await using var conn = new SqlConnection(BuildConnectionString(sql));
        await conn.OpenAsync(ct);

        await using (var deleteCmd = new SqlCommand(
            "DELETE FROM dbo.GroupMembers WHERE GroupId = @groupId AND IsNested = 0", conn))
        {
            deleteCmd.Parameters.Add("@groupId", SqlDbType.UniqueIdentifier).Value = groupId;
            await deleteCmd.ExecuteNonQueryAsync(ct);
        }

        const string insertSql = @"
INSERT INTO dbo.GroupMembers (GroupId, MemberType, MemberId, DisplayName, PrimarySmtpAddress, UserPrincipalName, UserType, IsNested, Path, LastImportedAtUtc)
VALUES (@groupId, @memberType, @memberId, @displayName, @primarySmtpAddress, @userPrincipalName, @userType, 0, '', SYSUTCDATETIME());
";

        foreach (var o in liveMembers)
        {
            var recipientType = Str(o, "RecipientTypeDetails");
            var memberId = Str(o, "ExternalDirectoryObjectId");
            if (string.IsNullOrWhiteSpace(memberId))
                memberId = Str(o, "Guid");

            await using var insertCmd = new SqlCommand(insertSql, conn);
            insertCmd.Parameters.Add("@groupId", SqlDbType.UniqueIdentifier).Value = groupId;
            insertCmd.Parameters.Add("@memberType", SqlDbType.NVarChar, 64).Value = recipientType.Contains("Group", StringComparison.OrdinalIgnoreCase) ? "Group" : "User";
            insertCmd.Parameters.Add("@memberId", SqlDbType.NVarChar, 256).Value = memberId;
            insertCmd.Parameters.Add("@displayName", SqlDbType.NVarChar, 256).Value = Str(o, "DisplayName");
            insertCmd.Parameters.Add("@primarySmtpAddress", SqlDbType.NVarChar, 256).Value = Str(o, "PrimarySmtpAddress");
            insertCmd.Parameters.Add("@userPrincipalName", SqlDbType.NVarChar, 256).Value = Str(o, "UserPrincipalName");
            insertCmd.Parameters.Add("@userType", SqlDbType.NVarChar, 64).Value = recipientType;
            await insertCmd.ExecuteNonQueryAsync(ct);
        }
    }

    private async Task<IReadOnlyList<GroupMemberInfo>> GetMembersCachedAsync(DistributionGroupInfo group, CancellationToken ct = default)
    {
        // A group not sourced from the cache (e.g. a live-only search fallback) has no known
        // Groups.Id to look members up by.
        if (!Guid.TryParse(group.Name, out var groupId))
            return Array.Empty<GroupMemberInfo>();

        var sql = _settings.Current.Sql;
        await using var conn = new SqlConnection(BuildConnectionString(sql));
        await conn.OpenAsync(ct);

        var query = @"
SELECT
    MemberType,
    DisplayName,
    PrimarySmtpAddress,
    UserPrincipalName,
    UserType,
    IsNested,
    Path
FROM dbo.GroupMembers
WHERE GroupId = @groupId
ORDER BY DisplayName;
";

        await using var cmd = new SqlCommand(query, conn);
        cmd.Parameters.Add("@groupId", SqlDbType.UniqueIdentifier).Value = groupId;

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        var list = new List<GroupMemberInfo>();
        while (await reader.ReadAsync(ct))
        {
            list.Add(new GroupMemberInfo
            {
                DisplayName = reader.IsDBNull(1) ? "" : reader.GetString(1),
                PrimarySmtpAddress = reader.IsDBNull(2) ? "" : reader.GetString(2),
                UserPrincipalName = reader.IsDBNull(3) ? "" : reader.GetString(3),
                RecipientType = reader.IsDBNull(0) ? "" : reader.GetString(0),
                IsNested = !reader.IsDBNull(5) && reader.GetBoolean(5),
                Path = reader.IsDBNull(6) ? "" : reader.GetString(6),
            });
        }

        return list;
    }

    public async Task<IReadOnlyList<MemberOperationResult>> AddMembersAsync(DistributionGroupInfo group, IEnumerable<string> memberIdentities, CancellationToken ct = default)
    {
        var identity = string.IsNullOrWhiteSpace(group.PrimarySmtpAddress) ? group.Name : group.PrimarySmtpAddress;
        var results = new List<MemberOperationResult>();

        foreach (var raw in memberIdentities)
        {
            var resolved = ResolveMemberIdentity(raw);
            try
            {
                if (group.IsUnifiedGroup)
                {
                    // Team-connected M365 groups: Add-UnifiedGroupLinks intermittently fails with
                    // "We failed to update the group mailbox" because it conflicts with Teams'
                    // own provisioning sync. Microsoft's guidance is to change membership via
                    // Graph instead.
                    await _auth.Graph.Groups[ResolveGraphGroupId(group)].Members.Ref.PostAsync(
                        new Microsoft.Graph.Models.ReferenceCreate
                        {
                            OdataId = $"https://graph.microsoft.com/v1.0/users/{resolved}",
                        }, cancellationToken: ct);
                }
                else
                {
                    await _host.InvokeAsync(ps => ps
                        .AddCommand("Add-DistributionGroupMember")
                        .AddParameter("Identity", identity)
                        .AddParameter("Member", resolved)
                        .AddParameter("BypassSecurityGroupManagerCheck", true)
                        .AddParameter("ErrorAction", "Stop"), ct: ct);
                }
                results.Add(new MemberOperationResult(raw, true, null));
            }
            catch (Exception ex)
            {
                results.Add(new MemberOperationResult(raw, false, ex.Message));
            }
        }

        await SyncMembersAsync(group, ct);
        return results;
    }

    public async Task<IReadOnlyList<MemberOperationResult>> RemoveMembersAsync(DistributionGroupInfo group, IEnumerable<string> memberIdentities, CancellationToken ct = default)
    {
        var identity = string.IsNullOrWhiteSpace(group.PrimarySmtpAddress) ? group.Name : group.PrimarySmtpAddress;
        var results = new List<MemberOperationResult>();

        foreach (var raw in memberIdentities)
        {
            try
            {
                if (group.IsUnifiedGroup)
                {
                    var user = await _auth.Graph.Users[raw].GetAsync(cfg =>
                        cfg.QueryParameters.Select = new[] { "id" }, ct);
                    if (string.IsNullOrEmpty(user?.Id))
                        throw new InvalidOperationException($"Could not resolve '{raw}' to a directory user.");

                    await _auth.Graph.Groups[ResolveGraphGroupId(group)].Members[user.Id].Ref.DeleteAsync(cancellationToken: ct);
                }
                else
                {
                    await _host.InvokeAsync(ps => ps
                        .AddCommand("Remove-DistributionGroupMember")
                        .AddParameter("Identity", identity)
                        .AddParameter("Member", raw)
                        .AddParameter("BypassSecurityGroupManagerCheck", true)
                        .AddParameter("Confirm", false)
                        .AddParameter("ErrorAction", "Stop"), ct: ct);
                }
                results.Add(new MemberOperationResult(raw, true, null));
            }
            catch (Exception ex)
            {
                results.Add(new MemberOperationResult(raw, false, ex.Message));
            }
        }

        await SyncMembersAsync(group, ct);
        return results;
    }

    /// <summary>Cached groups already have their Graph id as Name; live lookups carry it separately.</summary>
    private static string ResolveGraphGroupId(DistributionGroupInfo group)
        => Guid.TryParse(group.Name, out _) ? group.Name : group.ExternalDirectoryObjectId;

    /// <summary>A bare SamAccountName (no "@") is completed with the configured default domain.</summary>
    private string ResolveMemberIdentity(string raw)
    {
        var value = (raw ?? "").Trim();
        if (value.Length == 0 || value.Contains('@'))
            return value;

        var domain = _settings.Current.M365.Domain.Trim();
        return domain.Length == 0 ? value : $"{value}@{domain}";
    }

    // ----- mapping helpers -----

    private static string BuildConnectionString(SqlSettings sql)
    {
        var builder = new SqlConnectionStringBuilder
        {
            DataSource = sql.Server,
            InitialCatalog = sql.Database,
            Encrypt = true,
            TrustServerCertificate = true,
            ConnectTimeout = 10,
            ApplicationName = "M365Manager",
        };

        if (sql.UseIntegratedSecurity)
            builder.IntegratedSecurity = true;
        else
        {
            builder.UserID = sql.UserId;
            builder.Password = sql.Password;
        }

        return builder.ConnectionString;
    }

    private static DistributionGroupInfo MapGroup(PSObject o) => new()
    {
        Name = Str(o, "Name"),
        DisplayName = Str(o, "DisplayName"),
        PrimarySmtpAddress = Str(o, "PrimarySmtpAddress"),
        Alias = Str(o, "Alias"),
        GroupType = Str(o, "GroupType"),
        RecipientTypeDetails = Str(o, "RecipientTypeDetails"),
        Notes = Str(o, "Notes"),
        ManagedBy = StrList(o, "ManagedBy"),
        ExternalDirectoryObjectId = Str(o, "ExternalDirectoryObjectId"),
    };

    private static string Str(PSObject o, string name)
        => o.Properties[name]?.Value?.ToString() ?? "";

    private static List<string> StrList(PSObject o, string name)
    {
        var value = o.Properties[name]?.Value;
        var list = new List<string>();

        if (value is IEnumerable seq and not string)
        {
            foreach (var item in seq)
                if (item is not null)
                    list.Add(item.ToString()!);
        }
        else if (value is not null)
        {
            list.Add(value.ToString()!);
        }

        return list;
    }
}
