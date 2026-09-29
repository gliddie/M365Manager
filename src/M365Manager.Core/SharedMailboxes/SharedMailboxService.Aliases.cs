using System.Collections;
using System.Data;
using System.Management.Automation;
using System.Text.RegularExpressions;
using M365Manager.Core.PowerShell;
using M365Manager.Data.Logging;
using Microsoft.Data.SqlClient;

namespace M365Manager.Core.SharedMailboxes;

/// <summary>
/// Secondary SMTP aliases - port of the "Add/Remove Email Alias from Shared Mailbox" option of
/// ShrMbxAdminMenu.ps1. That option was a Read-Host prompt with no ticket and no log; here it
/// carries a ticket number and writes to the audit log like everything else.
/// </summary>
public sealed partial class SharedMailboxService
{
    private const string AliasAction = "ChangeAliases";

    public async Task<SharedMailboxAddresses> GetAddressesAsync(string mailboxIdentity, CancellationToken ct = default)
    {
        var (mailbox, _) = await ResolveSharedMailboxForAliasesAsync(mailboxIdentity, ct);
        var (primary, aliases) = await ReadSmtpAddressesAsync(Str(mailbox, "ExchangeGuid"), ct);

        return new SharedMailboxAddresses
        {
            DisplayName = Str(mailbox, "DisplayName"),
            PrimarySmtpAddress = primary.Length > 0 ? primary : Str(mailbox, "PrimarySmtpAddress"),
            Aliases = aliases,
        };
    }

    public async Task<ChangeSharedMailboxAliasesResult> ChangeAliasesAsync(ChangeSharedMailboxAliasesRequest request, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        try
        {
            await LogAsync(correlationId, AliasAction, "STAR", "Alias change has started", request.TaskNumber, Severity.Info);

            if (string.IsNullOrWhiteSpace(request.TaskNumber))
                return await FailAliasesAsync(correlationId, request.TaskNumber, "Enter the ticket/task number authorizing this change.");

            var addresses = (request.Addresses ?? "")
                .Split(new[] { ',', ';', ' ', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(a => a.Trim().TrimStart('<').TrimEnd('>'))
                .Select(a => a.StartsWith("smtp:", StringComparison.OrdinalIgnoreCase) ? a[5..] : a)
                .Where(a => a.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (addresses.Count == 0)
                return await FailAliasesAsync(correlationId, request.TaskNumber, "Enter at least one e-mail address.");

            var invalid = addresses.Where(a => !SmtpPattern().IsMatch(a)).ToList();
            if (invalid.Count > 0)
                return await FailAliasesAsync(correlationId, request.TaskNumber, $"Not an e-mail address: {string.Join(", ", invalid)}.");

            if (request.MakePrimary && (request.Mode != AliasChangeMode.Add || addresses.Count != 1))
                return await FailAliasesAsync(correlationId, request.TaskNumber, "Only a single added address can be made the primary address.");

            PSObject mailbox;
            try
            {
                (mailbox, _) = await ResolveSharedMailboxForAliasesAsync(request.MailboxIdentity, ct);
            }
            catch (InvalidOperationException ex)
            {
                return await FailAliasesAsync(correlationId, request.TaskNumber, ex.Message);
            }

            var mailboxId = Str(mailbox, "ExchangeGuid");
            var target = Str(mailbox, "PrimarySmtpAddress");
            var (primary, aliases) = await ReadSmtpAddressesAsync(mailboxId, ct);

            await LogAsync(correlationId, AliasAction, "INPUT",
                $"Mailbox={target}; Mode={request.Mode}; Addresses={string.Join(", ", addresses)}; MakePrimary={request.MakePrimary}",
                request.TaskNumber, Severity.Info, target);

            var changed = new List<string>();
            var skipped = new List<string>();
            var failed = new List<string>();
            bool Has(string a) => string.Equals(a, primary, StringComparison.OrdinalIgnoreCase)
                                  || aliases.Contains(a, StringComparer.OrdinalIgnoreCase);

            foreach (var address in addresses)
            {
                if (request.Mode == AliasChangeMode.Add)
                {
                    if (Has(address))
                    {
                        skipped.Add(address);
                        continue;
                    }

                    // Checked up front so the message names who has it - Exchange's own error only
                    // says "is already being used by the proxy addresses or LegacyExchangeDN".
                    if (await RecipientUsingAddressAsync(address, mailboxId, ct) is { } owner)
                    {
                        failed.Add($"{address}: already used by '{owner}'");
                        continue;
                    }
                }
                else
                {
                    if (string.Equals(address, primary, StringComparison.OrdinalIgnoreCase))
                    {
                        failed.Add($"{address}: it is the primary address - make another address primary first");
                        continue;
                    }
                    if (!Has(address))
                    {
                        skipped.Add(address);
                        continue;
                    }
                }

                // One address per call, lower-case "smtp:" so it goes in as a secondary: a bad
                // entry (a domain the tenant doesn't accept, say) fails alone.
                try
                {
                    var key = request.Mode == AliasChangeMode.Add ? "add" : "remove";
                    await _host.InvokeAsync(ps => ps
                        .AddCommand("Set-Mailbox")
                        .AddParameter("Identity", mailboxId)
                        .AddParameter("EmailAddresses", new Hashtable { [key] = $"smtp:{address}" })
                        .AddParameter("ErrorAction", "Stop"), ct: ct);

                    changed.Add(address);
                    await LogAsync(correlationId, AliasAction, request.Mode == AliasChangeMode.Add ? "ADD" : "REMO",
                        $"{(request.Mode == AliasChangeMode.Add ? "Added" : "Removed")} alias {address}", request.TaskNumber, Severity.Success, target);
                }
                catch (Exception ex)
                {
                    failed.Add($"{address}: {ex.Message}");
                    await LogAsync(correlationId, AliasAction, "WARN", $"{address}: {ex.Message}", request.TaskNumber, Severity.Warning, target);
                }
            }

            var newPrimary = primary;
            if (request.MakePrimary && failed.Count == 0)
            {
                var address = addresses[0];
                try
                {
                    // WindowsEmailAddress, not PrimarySmtpAddress - Exchange Online's Set-Mailbox has
                    // no such parameter. The previous primary stays behind as an alias.
                    await _host.InvokeAsync(ps => ps
                        .AddCommand("Set-Mailbox")
                        .AddParameter("Identity", mailboxId)
                        .AddParameter("WindowsEmailAddress", address)
                        .AddParameter("ErrorAction", "Stop"), ct: ct);

                    newPrimary = address;
                    changed.Add($"{address} (now primary)");
                    await LogAsync(correlationId, AliasAction, "PRIM",
                        $"Primary address changed from {primary} to {address}; {primary} kept as an alias", request.TaskNumber, Severity.Success, target);
                    await UpdatePrimaryAddressCacheAsync(mailboxId, address, ct);
                }
                catch (Exception ex)
                {
                    failed.Add($"{address} could not be made primary: {ex.Message}");
                    await LogAsync(correlationId, AliasAction, "WARN", $"Could not make {address} primary: {ex.Message}", request.TaskNumber, Severity.Warning, target);
                }
            }

            await LogAsync(correlationId, AliasAction, "DONE",
                $"Alias change has completed: {changed.Count} changed, {skipped.Count} unchanged, {failed.Count} failed",
                request.TaskNumber, failed.Count == 0 ? Severity.Success : Severity.Warning, target);

            return new ChangeSharedMailboxAliasesResult
            {
                Succeeded = true,
                Changed = changed,
                Skipped = skipped,
                Failed = failed,
                PrimarySmtpAddress = newPrimary,
                CorrelationId = correlationId,
            };
        }
        catch (Exception ex)
        {
            return await FailAliasesAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    /// <summary>Resolves the mailbox and refuses anything that is not a shared mailbox.</summary>
    private async Task<(PSObject Mailbox, string DisplayName)> ResolveSharedMailboxForAliasesAsync(string mailboxIdentity, CancellationToken ct)
    {
        var identity = (mailboxIdentity ?? "").Trim();
        if (identity.Length == 0)
            throw new InvalidOperationException("Enter the shared mailbox's display name or address.");

        var (mailbox, error) = await ResolveMailboxAsync(identity, ct);
        if (mailbox is null)
            throw new InvalidOperationException(error);

        var displayName = Str(mailbox, "DisplayName");
        var type = Str(mailbox, "RecipientTypeDetails");
        if (!string.Equals(type, "SharedMailbox", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException($"'{displayName}' is a {type}, not a shared mailbox.");

        return (mailbox, displayName);
    }

    /// <summary>Primary SMTP address and the secondary ones, from EmailAddresses ("SMTP:" = primary).</summary>
    private async Task<(string Primary, List<string> Aliases)> ReadSmtpAddressesAsync(string mailboxId, CancellationToken ct)
    {
        var all = await PsValues.ExpandProperty(_host, "Get-Mailbox", mailboxId, "EmailAddresses", ct);

        var primary = all.FirstOrDefault(a => a.StartsWith("SMTP:", StringComparison.Ordinal)) is { } p ? p[5..] : "";
        var aliases = all
            .Where(a => a.StartsWith("smtp:", StringComparison.Ordinal))
            .Select(a => a[5..])
            .OrderBy(a => a, StringComparer.OrdinalIgnoreCase)
            .ToList();
        return (primary, aliases);
    }

    /// <summary>Display name of another recipient already using <paramref name="address"/>, or null.</summary>
    private async Task<string?> RecipientUsingAddressAsync(string address, string mailboxId, CancellationToken ct)
    {
        var matches = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-Recipient")
            .AddParameter("Identity", address)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

        var other = matches.FirstOrDefault(m => !string.Equals(Str(m, "ExchangeGuid"), mailboxId, StringComparison.OrdinalIgnoreCase));
        return other is null ? null : Or(Str(other, "DisplayName"), address);
    }

    private async Task UpdatePrimaryAddressCacheAsync(string mailboxId, string address, CancellationToken ct)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs) || !Guid.TryParse(mailboxId, out var guid))
            return;

        try
        {
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync(ct);
            await using var cmd = new SqlCommand(
                "UPDATE dbo.SharedMailboxes SET PrimarySmtpAddress = @Address, LastSeenAtUtc = SYSUTCDATETIME() WHERE ExchangeGuid = @Id", conn);
            cmd.Parameters.Add("@Address", SqlDbType.NVarChar, 256).Value = address;
            cmd.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = guid;
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch
        {
            // Cache write is best-effort; the scheduled import will reconcile this row regardless.
        }
    }

    [GeneratedRegex(@"^[^@\s]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$")]
    private static partial Regex SmtpPattern();

    private async Task<ChangeSharedMailboxAliasesResult> FailAliasesAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, AliasAction, "FAIL", message, taskNumber, Severity.Error);
        return ChangeSharedMailboxAliasesResult.Error(correlationId, message);
    }
}
