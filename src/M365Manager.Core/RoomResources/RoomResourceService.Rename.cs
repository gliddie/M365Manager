using System.Collections;
using System.Data;
using System.Management.Automation;
using System.Text.RegularExpressions;
using M365Manager.Core.PowerShell;
using M365Manager.Data.Logging;
using Microsoft.Data.SqlClient;

namespace M365Manager.Core.RoomResources;

/// <summary>
/// Rename - port of RoomResourceRename.ps1. Kept in its own file because the rest of the service
/// was already long; it is the same class.
///
/// Three legacy bugs are deliberately not carried over:
/// - A restricted room's delegate group was only renamed when the SITE changed, although its name
///   contains the room name too - after a plain rename it kept pointing at the old name.
/// - A general-use room moving site had only the FullAccess of its site delegate group swapped;
///   the calendar's ResourceDelegates kept the old site's group, so the old site went on approving
///   out-of-policy requests.
/// - The mailbox was addressed by its display name, which changes halfway through. Everything here
///   addresses it by ExchangeGuid.
/// </summary>
public sealed partial class RoomResourceService
{
    private const string RenameAction = "Rename";

    public async Task<RoomRenameLookup> LookUpForRenameAsync(string mailboxIdentity, CancellationToken ct = default)
    {
        var identity = (mailboxIdentity ?? "").Trim();
        if (identity.Length == 0)
            throw new InvalidOperationException("Enter the room's address or display name.");

        var found = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-Mailbox")
            .AddParameter("Identity", identity)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);
        if (found.Count == 0)
            throw new InvalidOperationException($"No mailbox found for '{identity}'. Enter its address instead.");
        if (found.Count > 1)
            throw new InvalidOperationException($"'{identity}' matches {found.Count} mailboxes - enter the room's address instead.");

        var mailbox = found[0];
        var displayName = Str(mailbox, "DisplayName");
        var type = Str(mailbox, "RecipientTypeDetails");
        var kind = type switch
        {
            "RoomMailbox" => ResourceKind.Room,
            "EquipmentMailbox" => ResourceKind.Equipment,
            _ => throw new InvalidOperationException($"'{displayName}' is a {type}, not a room or equipment mailbox."),
        };

        if (!Guid.TryParse(Str(mailbox, "ExchangeGuid"), out var exchangeGuid))
            throw new InvalidOperationException($"'{displayName}' has no ExchangeGuid - it cannot be renamed safely from here.");

        var address = Str(mailbox, "PrimarySmtpAddress");
        var delegates = await PsValues.ExpandProperty(_host, "Get-CalendarProcessing", address, "ResourceDelegates", ct);
        var bookIn = await PsValues.ExpandProperty(_host, "Get-CalendarProcessing", address, "BookInPolicy", ct);

        // Same reading as UpdateDetailsAsync: a BookInPolicy means restricted. For general use,
        // a site-wide delegate group tells the two delegate variants apart.
        var accessModel = bookIn.Count > 0
            ? RoomAccessModel.Restricted
            : delegates.Count > 0 && !delegates.Any(IsSiteWideDelegateGroup)
                ? RoomAccessModel.GeneralUseCustomDelegates
                : RoomAccessModel.GeneralUseSiteDelegates;

        var timeZone = "";
        try
        {
            timeZone = (await PsValues.ExpandProperty(_host, "Get-MailboxRegionalConfiguration", address, "TimeZone", ct)).FirstOrDefault() ?? "";
        }
        catch
        {
            // Only used to decide whether the time zone needs changing; unknown just means "set it".
        }

        var alias = Str(mailbox, "Alias");
        var space = displayName.IndexOf(' ');

        return new RoomRenameLookup
        {
            ExchangeGuid = exchangeGuid,
            DisplayName = displayName,
            PrimarySmtpAddress = address,
            Alias = alias,
            Kind = kind,
            AccessModel = accessModel,
            SiteCode = (space > 0 ? displayName[..space] : displayName).ToUpperInvariant(),
            Capacity = int.TryParse(Str(mailbox, "ResourceCapacity"), out var capacity) ? capacity : null,
            Building = AddressSegment(alias, @"Bldg([^.]+)"),
            Floor = AddressSegment(alias, @"FLR([^.]+)") switch
            {
                "Grnd" => "Ground",
                var f => f,
            },
            TimeZone = timeZone,
            DelegateGroups = delegates,
            UsersGroups = bookIn,
            ExternalDirectoryObjectId = Str(mailbox, "ExternalDirectoryObjectId"),
            UserPrincipalName = Str(mailbox, "UserPrincipalName"),
        };
    }

    /// <summary>A segment of the naming convention's address, e.g. "Bldg1" -> "1". Empty when absent.</summary>
    private static string AddressSegment(string alias, string pattern)
    {
        var match = Regex.Match(alias, pattern);
        return match.Success ? match.Groups[1].Value : "";
    }

    public async Task<RoomRenamePlan> PlanRenameAsync(RoomRenameLookup current, RoomRenameRequest request, CancellationToken ct = default)
    {
        var errors = new List<string>();
        var notes = new List<string>();

        var names = await _naming.BuildAsync(
            request.RawName, current.Kind, current.AccessModel,
            request.Building, request.Floor,
            current.Kind == ResourceKind.Room ? request.Capacity : "",
            ResolveMailboxDomain(""), ResolveGroupDomain(""), ct);

        if (names.SiteCode.Length == 0 || names.DisplayName.Trim().Length <= names.SiteCode.Length)
            errors.Add("Enter the new name as \"<SITE> <Name>\", e.g. \"NBK Meeting Room 1\".");
        if (current.Kind == ResourceKind.Room && !int.TryParse(request.Capacity?.Trim(), out _))
            errors.Add("Room capacity must be a number.");

        var sameName = string.Equals(names.DisplayName, current.DisplayName, StringComparison.Ordinal);
        var sameAddress = string.Equals(names.Address, current.PrimarySmtpAddress, StringComparison.OrdinalIgnoreCase);
        if (sameName && sameAddress)
            errors.Add("The new name and address are the same as the current ones - nothing to rename.");

        // Legacy refused to run when the address or the alias was taken; a rename onto another
        // room's identity would fail halfway through, after the display name had already changed.
        if (!sameAddress && await RecipientOtherThanAsync(names.Address, current, ct) is { } addressOwner)
            errors.Add($"{names.Address} is already used by '{addressOwner}'.");
        if (!string.Equals(names.Alias, current.Alias, StringComparison.OrdinalIgnoreCase)
            && await RecipientOtherThanAsync(names.Alias, current, ct) is { } aliasOwner)
            errors.Add($"The alias {names.Alias} is already used by '{aliasOwner}'.");

        var siteChanged = !string.Equals(names.SiteCode, current.SiteCode, StringComparison.OrdinalIgnoreCase);
        var site = names.SiteCode.Length > 0 ? await _sites.FindByCodeAsync(names.SiteCode, ct) : null;
        if (siteChanged && site is null && names.SiteCode.Length > 0)
            errors.Add($"Site '{names.SiteCode}' is not maintained yet - add it under Settings > Room & resource sites (time zone + regional admin group).");

        // The room's own groups follow its name; site-wide delegate groups are shared and are
        // swapped, never renamed.
        var renames = new List<(string From, string To)>();
        (string From, string To)? delegateSwap = null;

        if (current.AccessModel == RoomAccessModel.GeneralUseSiteDelegates)
        {
            if (siteChanged)
            {
                var from = current.DelegateGroups.FirstOrDefault(IsSiteWideDelegateGroup) ?? "";
                delegateSwap = (from, names.DelegateGroup);
            }
        }
        else
        {
            var ownDelegate = current.DelegateGroups.FirstOrDefault(g => !IsSiteWideDelegateGroup(g));
            if (ownDelegate is null)
                notes.Add("The room has no delegate group of its own - nothing to rename there.");
            else if (!string.Equals(ownDelegate, names.DelegateGroup, StringComparison.OrdinalIgnoreCase))
                renames.Add((ownDelegate, names.DelegateGroup));
        }

        if (current.AccessModel == RoomAccessModel.Restricted && names.UsersGroup.Length > 0)
        {
            var ownUsers = current.UsersGroups.FirstOrDefault();
            if (ownUsers is not null && !string.Equals(ownUsers, names.UsersGroup, StringComparison.OrdinalIgnoreCase))
                renames.Add((ownUsers, names.UsersGroup));
        }

        foreach (var (from, to) in renames)
        {
            if (await TryProbeAsync("Get-DistributionGroup", to, ct) == true)
                errors.Add($"A group named {to} already exists, so {from} cannot be renamed to it.");
        }

        (string From, string To)? roomListMove = null;
        if (siteChanged && current.Kind == ResourceKind.Room)
        {
            var oldList = current.AccessModel == RoomAccessModel.Restricted
                ? $"{current.SiteCode} Restricted Rooms"
                : $"{current.SiteCode} Conference Rooms";
            roomListMove = (oldList, names.RoomListName);
        }

        var newTimeZone = site?.TimeZone is { Length: > 0 } tz && !string.Equals(tz, current.TimeZone, StringComparison.OrdinalIgnoreCase)
            ? tz
            : null;

        if (!sameAddress)
            notes.Add($"The previous address {current.PrimarySmtpAddress} is kept as an alias.");
        if (current.UserPrincipalName.Length > 0 && !string.Equals(current.UserPrincipalName, names.Address, StringComparison.OrdinalIgnoreCase))
            notes.Add($"The sign-in name (UPN) moves from {current.UserPrincipalName} to {names.Address}.");

        return new RoomRenamePlan
        {
            Current = current,
            NewNames = names,
            SiteChanged = siteChanged,
            GroupRenames = renames,
            DelegateSwap = delegateSwap,
            RoomListMove = roomListMove,
            NewTimeZone = newTimeZone,
            NewRegionalAdminGroup = siteChanged && site?.RegionalAdminGroup is { Length: > 0 } admins ? admins : null,
            Errors = errors,
            Notes = notes,
        };
    }

    public async Task<RoomRenameResult> RenameAsync(RoomRenameRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        try
        {
            await LogAsync(correlationId, RenameAction, "STAR", "Room/resource rename has started", request.TaskNumber, Severity.Info);

            if (string.IsNullOrWhiteSpace(request.TaskNumber))
                return await FailRenameAsync(correlationId, request.TaskNumber, "Enter the ticket/task number authorizing this rename.");

            // Looked up and planned again rather than trusting the form: the rename has to act on
            // the room as it is now.
            Progress(onProgress, "Reading the room...");
            RoomRenameLookup current;
            try
            {
                current = await LookUpForRenameAsync(request.MailboxIdentity, ct);
            }
            catch (InvalidOperationException ex)
            {
                return await FailRenameAsync(correlationId, request.TaskNumber, ex.Message);
            }

            var plan = await PlanRenameAsync(current, request, ct);
            if (plan.Errors.Count > 0)
                return await FailRenameAsync(correlationId, request.TaskNumber, string.Join(" ", plan.Errors));

            var names = plan.NewNames;
            var mailboxId = current.ExchangeGuid.ToString();
            var addressChanged = !string.Equals(names.Address, current.PrimarySmtpAddress, StringComparison.OrdinalIgnoreCase);

            await LogAsync(correlationId, RenameAction, "INPUT",
                $"{current.DisplayName} ({current.PrimarySmtpAddress}) -> {names.DisplayName} ({names.Address}); " +
                $"Access={current.AccessModel}; SiteChanged={plan.SiteChanged}; Delegates={string.Join(", ", current.DelegateGroups)}; " +
                $"Users={string.Join(", ", current.UsersGroups)}; TimeZone={current.TimeZone}", request.TaskNumber, Severity.Info);

            // 1. The mailbox. Address added first and promoted in a second call, so the old one
            //    stays behind as an alias - same two-step as the shared mailbox rename.
            Progress(onProgress, $"Renaming to {names.DisplayName} ({names.Address})...");
            var hasNewAddress = (await PsValues.ExpandProperty(_host, "Get-Mailbox", mailboxId, "EmailAddresses", ct))
                .Any(a => string.Equals(a.Contains(':') ? a[(a.IndexOf(':') + 1)..] : a, names.Address, StringComparison.OrdinalIgnoreCase));

            await _host.InvokeAsync(ps =>
            {
                ps.AddCommand("Set-Mailbox")
                  .AddParameter("Identity", mailboxId)
                  .AddParameter("Name", names.DisplayName)
                  .AddParameter("DisplayName", names.DisplayName)
                  .AddParameter("Alias", names.Alias);
                if (current.Kind == ResourceKind.Room && int.TryParse(request.Capacity?.Trim(), out var capacity))
                    ps.AddParameter("ResourceCapacity", capacity);
                if (!hasNewAddress)
                    ps.AddParameter("EmailAddresses", new Hashtable { ["add"] = names.Address });
                ps.AddParameter("ErrorAction", "Stop");
            }, ct: ct);

            if (addressChanged)
            {
                // Exchange Online's Set-Mailbox has no -PrimarySmtpAddress (that is on-premises and
                // Set-DistributionGroup). For a cloud mailbox, WindowsEmailAddress sets the primary
                // SMTP address and keeps the previous one as a proxy.
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-Mailbox")
                    .AddParameter("Identity", mailboxId)
                    .AddParameter("WindowsEmailAddress", names.Address)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
            }

            await LogAsync(correlationId, RenameAction, "RENA",
                $"Renamed {current.DisplayName} ({current.PrimarySmtpAddress}) to {names.DisplayName} ({names.Address})"
                + (addressChanged ? $"; previous address kept as an alias" : ""), request.TaskNumber, Severity.Success);

            // Everything after this point is a follow-up: the room is renamed and usable, so a
            // failing step is reported and the rest still runs.
            var changes = new List<string>();
            var warnings = new List<string>();

            async Task StepAsync(string label, Func<Task<string?>> step)
            {
                try
                {
                    if (await step() is { } done)
                    {
                        changes.Add(done);
                        await LogAsync(correlationId, RenameAction, "CHNG", done, request.TaskNumber, Severity.Success);
                    }
                }
                catch (Exception ex)
                {
                    var warning = $"{label} failed: {ex.Message}";
                    warnings.Add(warning);
                    await LogAsync(correlationId, RenameAction, "WARN", warning, request.TaskNumber, Severity.Warning);
                    Progress(onProgress, $"WARNING: {warning}");
                }
            }

            // 2. UPN, so sign-in name and address keep matching (legacy Update-MgUser).
            if (Guid.TryParse(current.ExternalDirectoryObjectId, out _)
                && !string.Equals(current.UserPrincipalName, names.Address, StringComparison.OrdinalIgnoreCase))
            {
                await StepAsync("Changing the sign-in name (UPN)", async () =>
                {
                    Progress(onProgress, "Changing the sign-in name (UPN)...");
                    using var _ = await _graph.PatchAsync($"/users/{current.ExternalDirectoryObjectId}",
                        new { userPrincipalName = names.Address }, ct);
                    return $"UPN changed from {current.UserPrincipalName} to {names.Address}";
                });
            }

            // 3. The room's own groups. The calendar and the permissions point at the group
            //    objects, not their names, so renaming them in place keeps every reference intact.
            foreach (var (from, to) in plan.GroupRenames)
            {
                await StepAsync($"Renaming group {from}", async () =>
                {
                    Progress(onProgress, $"Renaming group {from}...");
                    await RenameRoomGroupAsync(from, to, ct);
                    return $"Group {from} renamed to {to}";
                });
            }

            // 4. General-use room moving site: swap the shared site delegate group - in the
            //    calendar as well as in the mailbox permissions.
            if (plan.DelegateSwap is { } swap)
            {
                await StepAsync("Switching the site delegate group", async () =>
                {
                    Progress(onProgress, $"Switching delegates to {swap.To}...");
                    var exists = await TryProbeAsync("Get-DistributionGroup", swap.To, ct) == true;
                    if (!exists)
                    {
                        await EnsureAccessGroupAsync(correlationId, request.TaskNumber, swap.To,
                            $"{swap.To}@{ResolveGroupDomain("")}", false, "", "Room Delegates", ct);
                    }

                    var delegates = current.DelegateGroups
                        .Where(d => !string.Equals(d, swap.From, StringComparison.OrdinalIgnoreCase))
                        .Append(swap.To)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToArray();

                    await _host.InvokeAsync(ps => ps
                        .AddCommand("Set-CalendarProcessing")
                        .AddParameter("Identity", mailboxId)
                        .AddParameter("ResourceDelegates", delegates)
                        .AddParameter("ErrorAction", "Stop"), ct: ct);

                    if (swap.From.Length > 0)
                        await RemoveFullAccessAsync(mailboxId, swap.From, ct);
                    await GrantFullAccessAsync(correlationId, request.TaskNumber, mailboxId, swap.To, "site delegates", ct);

                    return swap.From.Length > 0
                        ? $"Delegates switched from {swap.From} to {swap.To}"
                        : $"Delegates set to {swap.To}";
                });
            }

            // 5. Room list.
            if (plan.RoomListMove is { } move)
            {
                await StepAsync("Moving the room list membership", async () =>
                {
                    Progress(onProgress, $"Moving from room list {move.From} to {move.To}...");
                    try
                    {
                        await _host.InvokeAsync(ps => ps
                            .AddCommand("Remove-DistributionGroupMember")
                            .AddParameter("Identity", move.From)
                            .AddParameter("Member", mailboxId)
                            .AddParameter("BypassSecurityGroupManagerCheck", true)
                            .AddParameter("Confirm", false)
                            .AddParameter("ErrorAction", "Stop"), ct: ct);
                    }
                    catch (Exception ex)
                    {
                        // Not being on the old list (or the list not existing) is fine - the room
                        // still has to end up on the new one.
                        await LogAsync(correlationId, RenameAction, "INFO", $"Not removed from {move.From}: {ex.Message}", request.TaskNumber, Severity.Info);
                    }

                    if (await TryProbeAsync("Get-DistributionGroup", move.To, ct) != true)
                    {
                        await _host.InvokeAsync(ps =>
                        {
                            ps.AddCommand("New-DistributionGroup")
                              .AddParameter("Name", move.To)
                              .AddParameter("RoomList", true)
                              .AddParameter("ErrorAction", "Stop");
                            AddManagedBy(ps, _settings.Current.RoomResources.RoomListOwnerGroup);
                        }, ct: ct);
                        await LogAsync(correlationId, RenameAction, "NEW", $"Created room list {move.To}", request.TaskNumber, Severity.Success);
                    }

                    await _host.InvokeAsync(ps => ps
                        .AddCommand("Add-DistributionGroupMember")
                        .AddParameter("Identity", move.To)
                        .AddParameter("Member", mailboxId)
                        .AddParameter("BypassSecurityGroupManagerCheck", true)
                        .AddParameter("ErrorAction", "Stop"), ct: ct);

                    return $"Room list {move.From} -> {move.To}";
                });
            }

            // 6. Time zone of the (new) site.
            if (plan.NewTimeZone is { } timeZone)
            {
                await StepAsync("Setting the time zone", async () =>
                {
                    Progress(onProgress, $"Setting time zone {timeZone}...");
                    await ApplyTimeZoneAsync(mailboxId, timeZone, ct);
                    return $"Time zone {(current.TimeZone.Length > 0 ? current.TimeZone + " -> " : "")}{timeZone}";
                });
            }

            // 7. Regional RRS admins: the new site's group in, other regional ones out. The global
            //    admin group is not "MBX.*RRS.Admins" and is left alone, as in the legacy script.
            if (plan.NewRegionalAdminGroup is { } regional)
            {
                await StepAsync("Switching the regional admin group", async () =>
                {
                    Progress(onProgress, $"Switching regional admins to {regional}...");
                    var perms = await _host.InvokeAsync(ps => ps
                        .AddCommand("Get-MailboxPermission")
                        .AddParameter("Identity", mailboxId)
                        .AddParameter("ErrorAction", "Stop"), ct: ct);

                    var removed = new List<string>();
                    var hasNew = false;
                    foreach (var user in perms.Select(p => Str(p, "User").Split('/').Last().Trim()).Distinct(StringComparer.OrdinalIgnoreCase))
                    {
                        if (string.Equals(user, regional, StringComparison.OrdinalIgnoreCase))
                        {
                            hasNew = true;
                            continue;
                        }
                        if (Regex.IsMatch(user, @"^MBX\..*RRS\.Admins$", RegexOptions.IgnoreCase))
                        {
                            await RemoveFullAccessAsync(mailboxId, user, ct);
                            removed.Add(user);
                        }
                    }
                    if (!hasNew)
                        await GrantFullAccessAsync(correlationId, request.TaskNumber, mailboxId, regional, "regional RRS admins", ct);

                    return removed.Count > 0
                        ? $"Regional admins {string.Join(", ", removed)} -> {regional}"
                        : $"Regional admins set to {regional}";
                });
            }

            await UpdateRenameCacheAsync(current.ExchangeGuid, plan, request, ct);

            string? mailWarning = null;
            if (!string.IsNullOrWhiteSpace(request.RequesterIdentity))
            {
                Progress(onProgress, $"Sending confirmation e-mail as {_mail.SenderDescription}...");
                mailWarning = await NotifyRequesterAsync(correlationId, RenameAction, request.RequesterIdentity, request.TaskNumber,
                    RoomNotificationMail.Renamed(current.DisplayName, names.DisplayName, current.PrimarySmtpAddress, names.Address,
                        plan.RoomListMove?.To, request.TaskNumber), ct);
                if (mailWarning is not null)
                    warnings.Add(mailWarning);
            }

            await LogAsync(correlationId, RenameAction, "DONE", "Room/resource rename has completed", request.TaskNumber, Severity.Success);
            Progress(onProgress, $"Renamed to '{names.DisplayName}'.");

            return new RoomRenameResult
            {
                Succeeded = true,
                PreviousDisplayName = current.DisplayName,
                PreviousPrimarySmtpAddress = current.PrimarySmtpAddress,
                DisplayName = names.DisplayName,
                PrimarySmtpAddress = names.Address,
                Changes = changes,
                WarningMessage = warnings.Count == 0 ? null : string.Join(" ", warnings),
                CorrelationId = correlationId,
            };
        }
        catch (Exception ex)
        {
            return await FailRenameAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    /// <summary>Display name of a recipient answering to <paramref name="identity"/> that is not this room, or null.</summary>
    private async Task<string?> RecipientOtherThanAsync(string identity, RoomRenameLookup current, CancellationToken ct)
    {
        var matches = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-Recipient")
            .AddParameter("Identity", identity)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

        var other = matches.FirstOrDefault(m =>
            !string.Equals(Str(m, "ExchangeGuid"), current.ExchangeGuid.ToString(), StringComparison.OrdinalIgnoreCase)
            && !string.Equals(Str(m, "PrimarySmtpAddress"), current.PrimarySmtpAddress, StringComparison.OrdinalIgnoreCase));
        return other is null ? null : Str(other, "DisplayName");
    }

    /// <summary>
    /// Renames one of the room's own groups in place: name, display name and alias, the new
    /// address added on the group's own domain and then promoted, so the old one stays as an alias.
    /// </summary>
    private async Task RenameRoomGroupAsync(string from, string to, CancellationToken ct)
    {
        var groups = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-DistributionGroup")
            .AddParameter("Identity", from)
            .AddParameter("ErrorAction", "Stop"), ct: ct);
        if (groups.Count == 0)
            throw new InvalidOperationException($"'{from}' no longer resolves to a group.");

        var group = groups[0];
        var groupId = Str(group, "Guid") is { Length: > 0 } guid ? guid : from;
        var currentAddress = Str(group, "PrimarySmtpAddress");
        var at = currentAddress.IndexOf('@');
        var newAddress = $"{to}@{(at > 0 ? currentAddress[(at + 1)..] : ResolveGroupDomain(""))}";

        await _host.InvokeAsync(ps => ps
            .AddCommand("Set-DistributionGroup")
            .AddParameter("Identity", groupId)
            .AddParameter("Name", to)
            .AddParameter("DisplayName", to)
            .AddParameter("Alias", to)
            .AddParameter("EmailAddresses", new Hashtable { ["add"] = newAddress })
            .AddParameter("BypassSecurityGroupManagerCheck", true)
            .AddParameter("ErrorAction", "Stop"), ct: ct);

        await _host.InvokeAsync(ps => ps
            .AddCommand("Set-DistributionGroup")
            .AddParameter("Identity", groupId)
            .AddParameter("PrimarySmtpAddress", newAddress)
            .AddParameter("BypassSecurityGroupManagerCheck", true)
            .AddParameter("ErrorAction", "Stop"), ct: ct);
    }

    private async Task RemoveFullAccessAsync(string mailboxId, string trustee, CancellationToken ct)
    {
        await _host.InvokeAsync(ps => ps
            .AddCommand("Remove-MailboxPermission")
            .AddParameter("Identity", mailboxId)
            .AddParameter("User", trustee)
            .AddParameter("AccessRights", "FullAccess")
            .AddParameter("Confirm", false)
            .AddParameter("ErrorAction", "Stop"), ct: ct);
    }

    /// <summary>Keyed on ExchangeGuid - the one identifier a rename does not change.</summary>
    private async Task UpdateRenameCacheAsync(Guid exchangeGuid, RoomRenamePlan plan, RoomRenameRequest request, CancellationToken ct)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return;

        try
        {
            var names = plan.NewNames;
            var delegateGroup = plan.DelegateSwap?.To
                                ?? plan.GroupRenames.FirstOrDefault(r => r.To == names.DelegateGroup).To;
            var usersGroup = plan.GroupRenames.FirstOrDefault(r => r.To == names.UsersGroup && names.UsersGroup.Length > 0).To;

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync(ct);

            const string sql = """
                UPDATE dbo.RoomResources SET
                    DisplayName = @DisplayName, Alias = @Alias, PrimarySmtpAddress = @Address,
                    Capacity = COALESCE(@Capacity, Capacity),
                    Building = @Building, Floor = @Floor, SiteCode = @SiteCode,
                    TimeZone = COALESCE(@TimeZone, TimeZone),
                    RoomListName = COALESCE(@RoomList, RoomListName),
                    DelegateGroup = COALESCE(@DelegateGroup, DelegateGroup),
                    UsersGroup = COALESCE(@UsersGroup, UsersGroup),
                    LastSeenAtUtc = SYSUTCDATETIME()
                WHERE ExchangeGuid = @Id;
                """;

            await using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = exchangeGuid;
            cmd.Parameters.Add("@DisplayName", SqlDbType.NVarChar, 256).Value = names.DisplayName;
            cmd.Parameters.Add("@Alias", SqlDbType.NVarChar, 128).Value = names.Alias;
            cmd.Parameters.Add("@Address", SqlDbType.NVarChar, 256).Value = names.Address;
            cmd.Parameters.Add("@Capacity", SqlDbType.Int).Value =
                int.TryParse(request.Capacity?.Trim(), out var cap) ? cap : DBNull.Value;
            cmd.Parameters.Add("@Building", SqlDbType.NVarChar, 128).Value = NullableParam(request.Building);
            cmd.Parameters.Add("@Floor", SqlDbType.NVarChar, 32).Value = NullableParam(request.Floor);
            cmd.Parameters.Add("@SiteCode", SqlDbType.NVarChar, 16).Value = NullableParam(names.SiteCode);
            cmd.Parameters.Add("@TimeZone", SqlDbType.NVarChar, 128).Value = NullableParam(plan.NewTimeZone);
            cmd.Parameters.Add("@RoomList", SqlDbType.NVarChar, 256).Value = NullableParam(plan.RoomListMove?.To);
            cmd.Parameters.Add("@DelegateGroup", SqlDbType.NVarChar, 256).Value = NullableParam(delegateGroup);
            cmd.Parameters.Add("@UsersGroup", SqlDbType.NVarChar, 256).Value = NullableParam(usersGroup);
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch
        {
            // Cache write is best-effort; the scheduled import will reconcile this row regardless.
        }
    }

    private async Task<RoomRenameResult> FailRenameAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, RenameAction, "FAIL", message, taskNumber, Severity.Error);
        return RoomRenameResult.Failed(correlationId, message);
    }
}
