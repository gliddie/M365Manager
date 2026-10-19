using System.Collections;
using System.Management.Automation;
using M365Manager.Core.PowerShell;
using M365Manager.Data.Logging;

namespace M365Manager.Core.RoomResources;

/// <summary>
/// Capacity change, part of UpdateDetailsAsync. The capacity is the last segment of a room's
/// address ("BRM.Flr2.MeetingRoom.10@..."), so changing only ResourceCapacity left the address,
/// alias and sign-in name naming the old number (2026-10-19). Only that segment is swapped - see
/// <see cref="RoomNamingService.WithCapacity"/> for why this is not a rename.
/// </summary>
public sealed partial class RoomResourceService
{
    private sealed record CapacityChange(
        string? Error,
        string? NewAddress,
        string? NewAlias,
        IReadOnlyList<string> Changes,
        IReadOnlyList<string> ReadableChanges,
        IReadOnlyList<string> Warnings);

    private async Task<CapacityChange> ChangeCapacityAsync(
        Guid correlationId, string taskNumber, PSObject mailbox, int? currentCapacity, int capacity,
        Action<string>? onProgress, CancellationToken ct)
    {
        var address = Str(mailbox, "PrimarySmtpAddress");
        var alias = Str(mailbox, "Alias");
        var upn = Str(mailbox, "UserPrincipalName");
        var exchangeGuid = Str(mailbox, "ExchangeGuid");
        var mailboxId = Guid.TryParse(exchangeGuid, out _) ? exchangeGuid : address;

        var newAddress = RoomNamingService.WithCapacity(address, currentCapacity, capacity);
        var newAlias = RoomNamingService.WithCapacity(alias, currentCapacity, capacity);
        var newUpn = RoomNamingService.WithCapacity(upn, currentCapacity, capacity);

        // Checked before anything changes: a taken address would fail halfway, after the
        // capacity had already been set.
        if (newAddress is not null && await RecipientOtherThanAsync(newAddress, exchangeGuid, address, ct) is { } addressOwner)
            return Fail($"{newAddress} is already used by '{addressOwner}' - the capacity was not changed.");
        if (newAlias is not null && await RecipientOtherThanAsync(newAlias, exchangeGuid, address, ct) is { } aliasOwner)
            return Fail($"The alias {newAlias} is already used by '{aliasOwner}' - the capacity was not changed.");

        var changes = new List<string>();
        var readable = new List<string>();
        var warnings = new List<string>();

        Progress(onProgress, newAddress is null ? "Updating capacity..." : $"Updating capacity and address to {newAddress}...");

        // Address added first and promoted in a second call, so the old one stays behind as an
        // alias - the same two steps as the rename.
        var hasNewAddress = newAddress is not null
            && (await PsValues.ExpandProperty(_host, "Get-Mailbox", mailboxId, "EmailAddresses", ct))
                .Any(a => string.Equals(a.Contains(':') ? a[(a.IndexOf(':') + 1)..] : a, newAddress, StringComparison.OrdinalIgnoreCase));

        await _host.InvokeAsync(ps =>
        {
            ps.AddCommand("Set-Mailbox")
              .AddParameter("Identity", mailboxId)
              .AddParameter("ResourceCapacity", capacity);
            if (newAlias is not null)
                ps.AddParameter("Alias", newAlias);
            if (newAddress is not null && !hasNewAddress)
                ps.AddParameter("EmailAddresses", new Hashtable { ["add"] = newAddress });
            ps.AddParameter("ErrorAction", "Stop");
        }, ct: ct);

        changes.Add($"capacity={currentCapacity?.ToString() ?? "(none)"}->{capacity}");
        readable.Add($"The capacity is now {capacity} people.");
        if (newAlias is not null)
            changes.Add($"alias={alias}->{newAlias}");

        if (newAddress is not null)
        {
            // Exchange Online's Set-Mailbox has no -PrimarySmtpAddress; WindowsEmailAddress sets the
            // primary address and keeps the previous one as a proxy.
            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-Mailbox")
                .AddParameter("Identity", mailboxId)
                .AddParameter("WindowsEmailAddress", newAddress)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            changes.Add($"address={address}->{newAddress} (previous kept as alias)");
            readable.Add($"The room's e-mail address is now {newAddress}. The previous address {address} keeps working.");
        }
        else
        {
            warnings.Add($"The address {address} does not end in the old capacity ({currentCapacity?.ToString() ?? "none set"}), so only the capacity was changed - rename the room if the address should carry it.");
        }

        await LogAsync(correlationId, "UpdateDetails", "CAPA", string.Join("; ", changes), taskNumber, Severity.Success);

        // The room is usable under its new address from here on, so a failing UPN change is a
        // warning, not a failure - same as in the rename.
        var objectId = Str(mailbox, "ExternalDirectoryObjectId");
        if (newUpn is not null && Guid.TryParse(objectId, out _))
        {
            try
            {
                Progress(onProgress, "Changing the sign-in name (UPN)...");
                using var _ = await _graph.PatchAsync($"/users/{objectId}", new { userPrincipalName = newUpn }, ct);
                var done = $"UPN changed from {upn} to {newUpn}";
                changes.Add(done);
                await LogAsync(correlationId, "UpdateDetails", "CHNG", done, taskNumber, Severity.Success);
            }
            catch (Exception ex)
            {
                var warning = $"Changing the sign-in name (UPN) to {newUpn} failed: {ex.Message}";
                warnings.Add(warning);
                await LogAsync(correlationId, "UpdateDetails", "WARN", warning, taskNumber, Severity.Warning);
            }
        }
        else if (newAddress is not null && upn.Length > 0)
        {
            warnings.Add($"The sign-in name {upn} does not end in the old capacity and was left as it is.");
        }

        return new CapacityChange(null, newAddress, newAlias, changes, readable, warnings);

        static CapacityChange Fail(string error) => new(error, null, null, [], [], []);
    }

    /// <summary>Display name of a recipient other than this room that answers to the identity; null when free.</summary>
    private async Task<string?> RecipientOtherThanAsync(string identity, string exchangeGuid, string address, CancellationToken ct)
    {
        var matches = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-Recipient")
            .AddParameter("Identity", identity)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

        var other = matches.FirstOrDefault(m =>
            !string.Equals(Str(m, "ExchangeGuid"), exchangeGuid, StringComparison.OrdinalIgnoreCase)
            && !string.Equals(Str(m, "PrimarySmtpAddress"), address, StringComparison.OrdinalIgnoreCase));
        return other is null ? null : Str(other, "DisplayName");
    }
}
