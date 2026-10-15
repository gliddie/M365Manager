using System.Security;
using System.Security.Cryptography;
using System.Text.Json;
using M365Manager.Core.M365;
using M365Manager.Core.PowerShell;

namespace M365Manager.Core.RoomResources;

/// <summary>
/// The Entra user behind a room mailbox. New-Mailbox -Room creates it disabled and without a usable
/// password; a Teams Rooms device needs it enabled with one that never expires (user request
/// 2026-10-15). Enabling and the password go through Exchange (Set-Mailbox
/// -EnableRoomMailboxAccount), the never-expire flag through Graph - see EnableWithPasswordAsync.
/// </summary>
public interface IRoomAccountService
{
    /// <summary>
    /// The room's Entra user ($select id, displayName, userPrincipalName, usageLocation, accountEnabled),
    /// or null. A room's address is normally its UPN, but a renamed or hand-made room can differ -
    /// then the mail attribute finds it.
    /// </summary>
    Task<JsonElement?> FindAsync(string address, CancellationToken ct = default);

    /// <summary>Polls <see cref="FindAsync"/> until the account shows up - right after New-Mailbox it may not yet.</summary>
    Task<JsonElement?> WaitForAsync(string address, TimeSpan wait, CancellationToken ct = default);

    /// <summary>
    /// Enables the account with a fresh random password and sets it to never expire. The caller
    /// shows the password once; it is never logged or stored. Throws when the account could not be
    /// enabled; a password that could not be set to never expire comes back as a warning.
    /// </summary>
    Task<RoomAccountPassword> EnableWithPasswordAsync(string roomAddress, string userPrincipalName, CancellationToken ct = default);
}

/// <param name="ExpiryWarning">Set when the password works but could not be set to never expire.</param>
public sealed record RoomAccountPassword(string Password, string? ExpiryWarning);

public sealed class RoomAccountService : IRoomAccountService
{
    private const string Select = "$select=id,displayName,userPrincipalName,usageLocation,accountEnabled";
    private static readonly TimeSpan Poll = TimeSpan.FromSeconds(15);

    private readonly GraphRestClient _graph;
    private readonly PowerShellHost _host;

    public RoomAccountService(GraphRestClient graph, PowerShellHost host)
    {
        _graph = graph;
        _host = host;
    }

    public async Task<JsonElement?> FindAsync(string address, CancellationToken ct = default)
    {
        try
        {
            using var doc = await _graph.GetAsync($"/users/{Uri.EscapeDataString(address)}?{Select}", ct: ct);
            return doc.RootElement.Clone();
        }
        catch
        {
            var filter = Uri.EscapeDataString($"mail eq '{address.Replace("'", "''")}'");
            var matches = await _graph.GetAllPagesAsync($"/users?$filter={filter}&{Select}", ct: ct);
            return matches.Count == 1 ? matches[0] : null;
        }
    }

    public async Task<JsonElement?> WaitForAsync(string address, TimeSpan wait, CancellationToken ct = default)
    {
        var deadline = DateTimeOffset.UtcNow + wait;
        while (true)
        {
            try
            {
                if (await FindAsync(address, ct) is { } found)
                    return found;
            }
            catch
            {
                // not there yet
            }

            if (DateTimeOffset.UtcNow >= deadline)
                return null;
            await Task.Delay(Poll, ct);
        }
    }

    public async Task<RoomAccountPassword> EnableWithPasswordAsync(string roomAddress, string userPrincipalName, CancellationToken ct = default)
    {
        var password = PasswordGenerator.Generate();

        // Through Exchange, not Graph: setting a password with Graph needs the delegated scope
        // User-PasswordProfile.ReadWrite.All, which Directory.ReadWrite.All does not include - the
        // first live run got 403 (2026-10-15). Exchange enables a room mailbox's account itself.
        // SecureString, so the console shows "(SecureString)" and never the password.
        using var secure = new SecureString();
        foreach (var ch in password)
            secure.AppendChar(ch);
        secure.MakeReadOnly();

        await _host.InvokeAsync(ps => ps
            .AddCommand("Set-Mailbox")
            .AddParameter("Identity", roomAddress)
            .AddParameter("EnableRoomMailboxAccount", true)
            .AddParameter("RoomMailboxPassword", secure)
            .AddParameter("Confirm", false)
            .AddParameter("ErrorAction", "Stop"), ct: ct);

        // Never expires: a Teams Rooms device cannot react to an expired password. No password in
        // this body. Not fatal - the account works, it just follows the tenant's expiry policy.
        string? warning = null;
        try
        {
            using var _ = await _graph.PatchAsync($"/users/{Uri.EscapeDataString(userPrincipalName)}",
                new { passwordPolicies = "DisablePasswordExpiration" }, ct);
        }
        catch (Exception ex)
        {
            warning = $"The password could not be set to never expire ({ex.Message}). Set it with "
                      + $"Update-MgUser -UserId {userPrincipalName} -PasswordPolicies DisablePasswordExpiration.";
        }

        return new RoomAccountPassword(password, warning);
    }
}

/// <summary>Passwords for room accounts. Ported from Switchboard (common area phone accounts).</summary>
public static class PasswordGenerator
{
    private const string Upper = "ABCDEFGHJKLMNPQRSTUVWXYZ";   // no I, O - easy to misread on a device keypad
    private const string Lower = "abcdefghijkmnopqrstuvwxyz";  // no l
    private const string Digits = "23456789";                  // no 0, 1
    private const string Special = "!#$%&*+-=?@_";

    /// <summary>
    /// 16 characters from a cryptographic random source, with at least one upper-case letter, one
    /// lower-case letter, one digit and one special character. Characters that are easily confused
    /// when typed on a device are left out.
    /// </summary>
    public static string Generate(int length = 16)
    {
        if (length < 4)
            throw new ArgumentOutOfRangeException(nameof(length));

        var all = Upper + Lower + Digits + Special;
        var chars = new List<char> { Pick(Upper), Pick(Lower), Pick(Digits), Pick(Special) };
        while (chars.Count < length)
            chars.Add(Pick(all));

        // Shuffle so the guaranteed classes are not always at the front.
        for (var i = chars.Count - 1; i > 0; i--)
        {
            var j = RandomNumberGenerator.GetInt32(i + 1);
            (chars[i], chars[j]) = (chars[j], chars[i]);
        }
        return new string(chars.ToArray());
    }

    private static char Pick(string set) => set[RandomNumberGenerator.GetInt32(set.Length)];
}
