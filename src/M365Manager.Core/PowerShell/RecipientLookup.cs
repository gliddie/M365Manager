namespace M365Manager.Core.PowerShell;

/// <summary>
/// Turning the identities an operator types (SamAccountName, UPN, e-mail) into the display names a
/// recipient of a notification e-mail can actually read. "118378, 115055" in a mail telling someone
/// who was given access is useless to them.
/// </summary>
public static class RecipientLookup
{
    /// <summary>
    /// Display name per identity, in the order given. An identity Exchange cannot resolve is kept
    /// as-is rather than dropped - a notification that silently omits someone would be worse than
    /// one naming them awkwardly.
    /// </summary>
    public static async Task<List<string>> ToDisplayNamesAsync(
        PowerShellHost host, IEnumerable<string> identities, CancellationToken ct = default)
    {
        var names = new List<string>();

        foreach (var identity in identities)
        {
            var value = (identity ?? "").Trim();
            if (value.Length == 0)
                continue;

            names.Add(await ToDisplayNameAsync(host, value, ct));
        }

        return names;
    }

    public static async Task<string> ToDisplayNameAsync(PowerShellHost host, string identity, CancellationToken ct = default)
    {
        try
        {
            var recipient = await host.InvokeAsync(ps => ps
                .AddCommand("Get-Recipient")
                .AddParameter("Identity", identity)
                .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

            if (recipient.Count > 0)
            {
                var display = recipient[0].Properties["DisplayName"]?.Value?.ToString() ?? "";
                if (display.Length > 0)
                    return display;
            }
        }
        catch
        {
            // Fall through - a lookup failure must not fail the operation the mail is about.
        }

        return identity;
    }
}
