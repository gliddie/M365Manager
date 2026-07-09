using System.Collections;
using System.Management.Automation;
using M365Manager.Core.M365;
using M365Manager.Core.PowerShell;

namespace M365Manager.Core.Exchange;

/// <summary>
/// Exchange Online operations via the hosted PowerShell runspace (bundled EXO module).
/// Connects with an access token from the signed-in admin, so actions run under their identity.
/// </summary>
public sealed class ExchangeService : IExchangeService
{
    private readonly PowerShellHost _host;
    private readonly IM365AuthService _auth;

    public ExchangeService(PowerShellHost host, IM365AuthService auth)
    {
        _host = host;
        _auth = auth;
    }

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

    public async Task<IReadOnlyList<GroupMemberInfo>> GetMembersAsync(DistributionGroupInfo group, CancellationToken ct = default)
    {
        var identity = string.IsNullOrWhiteSpace(group.PrimarySmtpAddress) ? group.Name : group.PrimarySmtpAddress;
        var isUnified = group.RecipientTypeDetails.Contains("GroupMailbox", StringComparison.OrdinalIgnoreCase);

        var results = isUnified
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

        return results.Select(o => new GroupMemberInfo
        {
            DisplayName = Str(o, "DisplayName"),
            PrimarySmtpAddress = Str(o, "PrimarySmtpAddress"),
            RecipientType = Str(o, "RecipientTypeDetails"),
        })
        .OrderBy(m => m.DisplayName)
        .ToList();
    }

    public async Task AddMemberAsync(DistributionGroupInfo group, string memberIdentity, CancellationToken ct = default)
    {
        var identity = string.IsNullOrWhiteSpace(group.PrimarySmtpAddress) ? group.Name : group.PrimarySmtpAddress;
        var isUnified = group.RecipientTypeDetails.Contains("GroupMailbox", StringComparison.OrdinalIgnoreCase);

        await _host.InvokeAsync(ps =>
        {
            if (isUnified)
            {
                ps.AddCommand("Add-UnifiedGroupLinks")
                  .AddParameter("Identity", identity)
                  .AddParameter("LinkType", "Members")
                  .AddParameter("Links", memberIdentity)
                  .AddParameter("ErrorAction", "Stop");
            }
            else
            {
                ps.AddCommand("Add-DistributionGroupMember")
                  .AddParameter("Identity", identity)
                  .AddParameter("Member", memberIdentity)
                  .AddParameter("BypassSecurityGroupManagerCheck", true)
                  .AddParameter("ErrorAction", "Stop");
            }
        }, ct: ct);
    }

    public async Task RemoveMemberAsync(DistributionGroupInfo group, string memberIdentity, CancellationToken ct = default)
    {
        var identity = string.IsNullOrWhiteSpace(group.PrimarySmtpAddress) ? group.Name : group.PrimarySmtpAddress;
        var isUnified = group.RecipientTypeDetails.Contains("GroupMailbox", StringComparison.OrdinalIgnoreCase);

        await _host.InvokeAsync(ps =>
        {
            if (isUnified)
            {
                ps.AddCommand("Remove-UnifiedGroupLinks")
                  .AddParameter("Identity", identity)
                  .AddParameter("LinkType", "Members")
                  .AddParameter("Links", memberIdentity)
                  .AddParameter("Confirm", false)
                  .AddParameter("ErrorAction", "Stop");
            }
            else
            {
                ps.AddCommand("Remove-DistributionGroupMember")
                  .AddParameter("Identity", identity)
                  .AddParameter("Member", memberIdentity)
                  .AddParameter("BypassSecurityGroupManagerCheck", true)
                  .AddParameter("Confirm", false)
                  .AddParameter("ErrorAction", "Stop");
            }
        }, ct: ct);
    }

    // ----- mapping helpers -----

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
