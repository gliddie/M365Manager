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
    private const string ExchangeScope = "https://outlook.office365.com/.default";

    private readonly PowerShellHost _host;
    private readonly IM365AuthService _auth;

    public ExchangeService(PowerShellHost host, IM365AuthService auth)
    {
        _host = host;
        _auth = auth;
    }

    public bool IsConnected { get; private set; }

    public async Task ConnectAsync(CancellationToken ct = default)
    {
        var user = _auth.CurrentUser
            ?? throw new InvalidOperationException("Sign in to M365 first (Settings → Sign in and test).");

        var organization = user.Upn.Contains('@')
            ? user.Upn[(user.Upn.IndexOf('@') + 1)..]
            : user.Upn;

        await _host.EnsureExchangeModuleAsync(ct);

        // Reuse the signed-in browser credential to get an Exchange token (no device code).
        var token = await _auth.GetAccessTokenAsync(ExchangeScope, ct);

        await _host.InvokeAsync(ps => ps
            .AddCommand("Connect-ExchangeOnline")
            .AddParameter("AccessToken", token)
            .AddParameter("Organization", organization)
            .AddParameter("ShowBanner", false)
            .AddParameter("ErrorAction", "Stop"), ct: ct);

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
