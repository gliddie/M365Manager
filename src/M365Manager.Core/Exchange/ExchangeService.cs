using System.Collections;
using System.Management.Automation;
using M365Manager.Core.M365;
using M365Manager.Core.PowerShell;

namespace M365Manager.Core.Exchange;

/// <summary>
/// Exchange Online operations via the hosted PowerShell runspace (bundled EXO module).
/// Connects interactively as the signed-in admin, so actions run under their identity.
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

    public async Task ConnectAsync(CancellationToken ct = default)
    {
        await _host.EnsureExchangeModuleAsync(ct);

        var upn = _auth.CurrentUser?.Upn;

        await _host.InvokeAsync(ps =>
        {
            ps.AddCommand("Connect-ExchangeOnline")
              .AddParameter("ShowBanner", false)
              .AddParameter("ErrorAction", "Stop");
            if (!string.IsNullOrWhiteSpace(upn))
                ps.AddParameter("UserPrincipalName", upn);
        }, ct);

        IsConnected = true;
    }

    public async Task<IReadOnlyList<DistributionGroupInfo>> SearchGroupsAsync(string search, CancellationToken ct = default)
    {
        var term = (search ?? "").Replace("'", "").Trim();
        var filter = $"DisplayName -like '*{term}*' -or Alias -like '*{term}*' -or PrimarySmtpAddress -like '*{term}*' -or Name -like '*{term}*'";

        var results = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-DistributionGroup")
            .AddParameter("Filter", filter)
            .AddParameter("ResultSize", 50), ct);

        return results.Select(MapGroup).ToList();
    }

    public async Task<DistributionGroupInfo?> GetGroupAsync(string identity, CancellationToken ct = default)
    {
        var results = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-DistributionGroup")
            .AddParameter("Identity", identity)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct);

        var first = results.FirstOrDefault();
        return first is null ? null : MapGroup(first);
    }

    public async Task<IReadOnlyList<GroupMemberInfo>> GetMembersAsync(string identity, CancellationToken ct = default)
    {
        var results = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-DistributionGroupMember")
            .AddParameter("Identity", identity)
            .AddParameter("ResultSize", "Unlimited")
            .AddParameter("ErrorAction", "Stop"), ct);

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
