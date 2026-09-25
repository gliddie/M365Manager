using System.Collections;
using System.Data;
using System.Management.Automation;
using System.Text;
using System.Text.Json;
using M365Manager.Core.PowerShell;
using M365Manager.Data.Logging;
using Microsoft.Data.SqlClient;

namespace M365Manager.Core.SharedMailboxes;

/// <summary>
/// Remove and recover - ports of ShrMbxRemove.ps1 and RecoverShrMbx.ps1. Kept apart from the
/// create/owner/rename half only because that file was already long; it is the same class.
/// </summary>
public sealed partial class SharedMailboxService
{
    private const string RemoveAction = "RemoveSharedMailbox";
    private const string RecoverAction = "RecoverSharedMailbox";

    /// <summary>Event code of the machine-readable log entry a recovery reads its prefill from.</summary>
    private const string RestoreDataCode = "RSTDATA";

    /// <summary>Everything a removal needs to know, gathered before anything is changed.</summary>
    private sealed record RemovalState(
        PSObject Mailbox,
        string DisplayName,
        string Address,
        Guid ExchangeGuid,
        List<(string Upn, string DisplayName, string Mail)> Owners,
        string Forwarding,
        List<SharedMailboxAccessGroupInfo> Groups,
        List<string> DirectAccess);

    public async Task<SharedMailboxRemovalPreview> PreviewRemovalAsync(string mailboxIdentity, CancellationToken ct = default)
    {
        var state = await GatherRemovalStateAsync(mailboxIdentity, ct);
        return new SharedMailboxRemovalPreview
        {
            DisplayName = state.DisplayName,
            PrimarySmtpAddress = state.Address,
            Owners = string.Join(", ", state.Owners.Select(o => o.DisplayName)),
            Forwarding = state.Forwarding,
            AccessGroups = state.Groups,
            DirectAccess = state.DirectAccess,
        };
    }

    public async Task<RemoveSharedMailboxResult> RemoveAsync(RemoveSharedMailboxRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, RemoveAction, "STAR", "Remove Shared Mailbox has started", request.TaskNumber, Severity.Info);

            if (string.IsNullOrWhiteSpace(request.TaskNumber))
                return await FailRemoveAsync(correlationId, request.TaskNumber, "Enter the ticket/task number authorizing this removal.");

            // Read fresh rather than trusting what the form showed: the removal acts on the state
            // of the mailbox now, and the snapshot has to describe exactly that.
            Progress(onProgress, "Looking the mailbox up...");
            RemovalState state;
            try
            {
                state = await GatherRemovalStateAsync(request.MailboxIdentity, ct);
            }
            catch (InvalidOperationException ex)
            {
                return await FailRemoveAsync(correlationId, request.TaskNumber, ex.Message);
            }

            await LogAsync(correlationId, RemoveAction, "INPUT",
                $"Mailbox={state.Address}; DeleteAccessGroups={request.DeleteAccessGroups}; " +
                $"Groups={string.Join(", ", state.Groups.Select(g => g.Name))}",
                request.TaskNumber, Severity.Info, state.Address);

            // Snapshot BEFORE anything is deleted - the same data the legacy script wrote to its
            // report file, into the audit log so it stays with the rest of the trail.
            Progress(onProgress, "Capturing pre-removal snapshot...");
            var snapshot = await BuildRemovalSnapshotAsync(state, ct);
            await LogAsync(correlationId, RemoveAction, "SNAP", snapshot, request.TaskNumber, Severity.Info, state.Address);

            var restoreData = new SharedMailboxRestoreData
            {
                ExchangeGuid = state.ExchangeGuid,
                DisplayName = state.DisplayName,
                PrimarySmtpAddress = state.Address,
                RemovedAtUtc = DateTime.UtcNow,
                TaskNumber = request.TaskNumber,
                Owners = state.Owners.Select(o => o.Upn).ToList(),
                Groups = state.Groups.Select(g => new SharedMailboxRestoreData.RestoreGroup
                {
                    Tier = g.Tier,
                    Name = g.Name,
                    Members = g.MemberIdentities.ToList(),
                }).ToList(),
            };
            await LogAsync(correlationId, RemoveAction, RestoreDataCode, JsonSerializer.Serialize(restoreData),
                request.TaskNumber, Severity.Info, state.ExchangeGuid.ToString());

            // The mailbox goes first, the groups after it - the reverse of the legacy order. If
            // Remove-Mailbox fails, nothing is lost yet; the other way round a failure left a
            // mailbox standing whose access groups were already gone for good.
            Progress(onProgress, "Removing the mailbox...");
            await _host.InvokeAsync(ps => ps
                .AddCommand("Remove-Mailbox")
                .AddParameter("Identity", state.Address)
                .AddParameter("Confirm", false)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            await LogAsync(correlationId, RemoveAction, "REMO", $"Removed mailbox {state.Address}", request.TaskNumber, Severity.Success, state.Address);

            var warnings = new List<string>();
            var removedGroups = new List<string>();
            if (request.DeleteAccessGroups)
            {
                foreach (var group in state.Groups)
                {
                    Progress(onProgress, $"Removing access group {group.Name}...");
                    try
                    {
                        await _host.InvokeAsync(ps => ps
                            .AddCommand("Remove-DistributionGroup")
                            .AddParameter("Identity", group.Address.Length > 0 ? group.Address : group.Name)
                            .AddParameter("BypassSecurityGroupManagerCheck", true)
                            .AddParameter("Confirm", false)
                            .AddParameter("ErrorAction", "Stop"), ct: ct);
                        removedGroups.Add(group.Name);
                        await LogAsync(correlationId, RemoveAction, "REMO", $"Removed access group {group.Name}", request.TaskNumber, Severity.Success, state.Address);
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"Access group {group.Name} could not be removed: {ex.Message}");
                        await LogAsync(correlationId, RemoveAction, "WARN", $"Could not remove access group {group.Name}: {ex.Message}", request.TaskNumber, Severity.Warning, state.Address);
                    }
                }
            }

            await MarkCacheDeletedAsync(state.ExchangeGuid, ct);

            var ownerMails = state.Owners.Select(o => MailOrUpn(o.Upn, o.Mail)).ToArray();
            if (ownerMails.Length == 0)
            {
                await LogAsync(correlationId, RemoveAction, "INFO", "No owner resolved - no confirmation e-mail sent", request.TaskNumber, Severity.Info, state.Address);
            }
            else
            {
                Progress(onProgress, $"Sending confirmation e-mail to the owner(s) as {_mail.SenderDescription}...");
                try
                {
                    await SendRemovalConfirmationMailAsync(ownerMails, state.DisplayName, removedGroups.Count > 0, request.TaskNumber, ct);
                    await LogAsync(correlationId, RemoveAction, "SEND",
                        $"Sent confirmation mail via {_mail.SenderDescription} to {string.Join(", ", ownerMails)}",
                        request.TaskNumber, Severity.Success, state.Address);
                }
                catch (Exception ex)
                {
                    warnings.Add($"The confirmation e-mail to {string.Join(", ", ownerMails)} could not be sent: {ex.Message}");
                    await LogAsync(correlationId, RemoveAction, "WARN", $"Could not send confirmation mail: {ex.Message}", request.TaskNumber, Severity.Warning, state.Address);
                }
            }

            await LogAsync(correlationId, RemoveAction, "DONE", "Remove Shared Mailbox has completed", request.TaskNumber, Severity.Success, state.Address);
            Progress(onProgress, $"'{state.DisplayName}' removed.");

            return new RemoveSharedMailboxResult
            {
                Succeeded = true,
                DisplayName = state.DisplayName,
                PrimarySmtpAddress = state.Address,
                RemovedGroups = removedGroups,
                CorrelationId = correlationId,
                WarningMessage = warnings.Count == 0 ? null : string.Join(" ", warnings),
            };
        }
        catch (Exception ex)
        {
            return await FailRemoveAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    public async Task<IReadOnlyList<DeletedSharedMailbox>> GetDeletedSharedMailboxesAsync(CancellationToken ct = default)
    {
        var results = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-Mailbox")
            .AddParameter("SoftDeletedMailbox", true)
            .AddParameter("RecipientTypeDetails", "SharedMailbox")
            .AddParameter("ResultSize", "Unlimited")
            .AddParameter("ErrorAction", "Stop"), ct: ct);

        return results
            .Select(m => (Mailbox: m, Guid: Guid.TryParse(Str(m, "ExchangeGuid"), out var g) ? g : Guid.Empty))
            .Where(x => x.Guid != Guid.Empty)
            .Select(x => new DeletedSharedMailbox
            {
                ExchangeGuid = x.Guid,
                DisplayName = Str(x.Mailbox, "DisplayName"),
                PrimarySmtpAddress = Str(x.Mailbox, "PrimarySmtpAddress"),
                WhenSoftDeleted = DateValue(x.Mailbox, "WhenSoftDeleted"),
            })
            .OrderByDescending(m => m.WhenSoftDeleted ?? DateTime.MinValue)
            .ToList();
    }

    public async Task<SharedMailboxRestoreData?> GetRestoreDataAsync(Guid exchangeGuid, CancellationToken ct = default)
    {
        try
        {
            var target = exchangeGuid.ToString();
            var entries = await _log.QueryAsync(new LogQuery
            {
                Area = "SharedMailboxes",
                Action = RemoveAction,
                SearchText = target,
                MaxResults = 50,
            }, ct);

            // Newest first, so a mailbox removed, recovered and removed again prefills from its
            // most recent removal.
            var entry = entries.FirstOrDefault(e =>
                e.EventCode == RestoreDataCode
                && string.Equals(e.TargetObject, target, StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(e.Message));

            return entry is null ? null : JsonSerializer.Deserialize<SharedMailboxRestoreData>(entry.Message!);
        }
        catch
        {
            // A prefill is a convenience. No SQL, an unreadable entry - the operator types instead.
            return null;
        }
    }

    public async Task<RecoverSharedMailboxResult> RecoverAsync(RecoverSharedMailboxRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        var target = request.ExchangeGuid.ToString();

        try
        {
            await LogAsync(correlationId, RecoverAction, "STAR", "Recover Shared Mailbox has started", request.TaskNumber, Severity.Info, target);

            if (string.IsNullOrWhiteSpace(request.TaskNumber))
                return await FailRecoverAsync(correlationId, request.TaskNumber, "Enter the ticket/task number authorizing this recovery.");

            // The legacy form refused to run without owners, and so does this: the recreated
            // groups need someone to manage them, and the .ED group needs at least one member.
            var owners = await ResolveOwnersAsync(request.OwnerIdentities, ct);
            if (owners.Count == 0)
                return await FailRecoverAsync(correlationId, request.TaskNumber, $"Could not resolve any owner from '{request.OwnerIdentities}' to a directory user.");
            var ownerUpns = owners.Select(o => o.Upn).ToArray();

            Progress(onProgress, "Looking up the deleted mailbox...");
            var deleted = await _host.InvokeAsync(ps => ps
                .AddCommand("Get-Mailbox")
                .AddParameter("SoftDeletedMailbox", true)
                .AddParameter("Identity", target)
                .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);
            if (deleted.Count == 0)
                return await FailRecoverAsync(correlationId, request.TaskNumber,
                    "The deleted mailbox was not found - it may have been restored already, or its 30-day recovery window has passed.");

            var displayName = Str(deleted[0], "DisplayName");
            var address = Str(deleted[0], "PrimarySmtpAddress");
            var atIndex = address.IndexOf('@');
            var domain = atIndex > 0 ? address[(atIndex + 1)..] : "";

            await LogAsync(correlationId, RecoverAction, "INPUT",
                $"Mailbox={displayName} ({address}); Owners={string.Join(", ", ownerUpns)}", request.TaskNumber, Severity.Info, target);

            Progress(onProgress, $"Restoring {displayName}...");
            await _host.InvokeAsync(ps => ps
                .AddCommand("Undo-SoftDeletedMailbox")
                .AddParameter("SoftDeletedObject", target)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            await LogAsync(correlationId, RecoverAction, "RSTR", $"Restored soft-deleted mailbox {displayName} ({address})", request.TaskNumber, Severity.Success, target);

            // Undo-SoftDeletedMailbox returns before the mailbox is usable again.
            PSObject? mailbox = null;
            for (var attempt = 1; attempt <= 20 && mailbox is null; attempt++)
            {
                var found = await _host.InvokeAsync(ps => ps
                    .AddCommand("Get-Mailbox")
                    .AddParameter("Identity", address)
                    .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);
                if (found.Count > 0)
                    mailbox = found[0];
                else
                {
                    Progress(onProgress, $"Waiting for the restored mailbox to become available... ({attempt}/20)");
                    await Task.Delay(TimeSpan.FromSeconds(6), ct);
                }
            }
            if (mailbox is null)
                return await FailRecoverAsync(correlationId, request.TaskNumber,
                    $"{displayName} was restored, but did not become available within two minutes, so its access groups were not recreated. " +
                    "Check it in a few minutes and set up its access groups by hand if needed.");

            await ApplyConfiguredPoliciesAsync(correlationId, RecoverAction, request.TaskNumber, address, onProgress, ct);

            var ownerMailTip = BuildOwnerMailTip(owners.Select(o => o.DisplayName).ToList());
            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-Mailbox")
                .AddParameter("Identity", address)
                .AddParameter("MailTip", ownerMailTip)
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            // Groups kept at removal time (the "also delete" box unticked) are still there and
            // still hold their permissions - reuse those rather than creating a duplicate.
            var existingGroups = await DiscoverAccessGroupsAsync(address, displayName, domain, ct);
            var tiers = await PlanTierMembersAsync(correlationId, RecoverAction, request.TaskNumber,
                ownerUpns, request.EditorMembers, request.AuthorMembers, request.ReaderMembers, ct);

            var warnings = new List<string>();
            var accessTiers = new List<AccessTier>();
            var groupLines = new List<string>();
            string? edGroupAddress = null, auGroupAddress = null, reGroupAddress = null;

            foreach (var (tier, membersCsv) in tiers)
            {
                try
                {
                    string groupName, groupAddress;
                    if (existingGroups.TryGetValue(tier, out var existing))
                    {
                        Progress(onProgress, $"Reusing existing .{tier} access group {existing}...");
                        groupName = existing;
                        groupAddress = (await PsValues.ExpandProperty(_host, "Get-DistributionGroup", existing, "PrimarySmtpAddress", ct)).FirstOrDefault() ?? existing;
                        await AddGroupMembersAsync(groupAddress, membersCsv, ct);
                        await LogAsync(correlationId, RecoverAction, "MEMBER", $"Reused {groupName}; added members: {membersCsv}", request.TaskNumber, Severity.Success, target);
                    }
                    else
                    {
                        (groupName, groupAddress) = await CreateAccessGroupAsync(correlationId, RecoverAction, request.TaskNumber,
                            displayName, address, domain, tier, membersCsv, ownerUpns, ownerMailTip, onProgress, ct);
                    }

                    accessTiers.Add(new AccessTier(tier, groupName, membersCsv));
                    groupLines.Add($".{tier} {groupName}");
                    switch (tier)
                    {
                        case "ED": edGroupAddress = groupAddress; break;
                        case "AU": auGroupAddress = groupAddress; break;
                        case "RE": reGroupAddress = groupAddress; break;
                    }
                }
                catch (Exception ex)
                {
                    // The mailbox is back either way; a missing tier is fixable by hand.
                    warnings.Add($"The .{tier} access group could not be set up: {ex.Message}");
                    await LogAsync(correlationId, RecoverAction, "WARN", $".{tier} access group failed: {ex.Message}", request.TaskNumber, Severity.Warning, target);
                }
            }

            var requireSenderAuth = string.Equals(Str(mailbox, "RequireSenderAuthenticationEnabled"), "True", StringComparison.OrdinalIgnoreCase);
            await UpsertCacheAsync(mailbox, displayName, Str(mailbox, "Alias"), address,
                edGroupAddress, auGroupAddress, reGroupAddress, string.Join("; ", ownerUpns), requireSenderAuth, ct);

            var ownerMails = owners.Select(o => MailOrUpn(o.Upn, o.Mail)).ToArray();
            Progress(onProgress, $"Sending confirmation e-mail to the owner(s) as {_mail.SenderDescription}...");
            try
            {
                await SendCreationConfirmationMailAsync(owners, displayName, address, request.TaskNumber,
                    !requireSenderAuth, accessTiers, ct, restored: true);
                await LogAsync(correlationId, RecoverAction, "SEND",
                    $"Sent confirmation mail via {_mail.SenderDescription} to {string.Join(", ", ownerMails)}",
                    request.TaskNumber, Severity.Success, target);
            }
            catch (Exception ex)
            {
                warnings.Add($"The confirmation e-mail to {string.Join(", ", ownerMails)} could not be sent: {ex.Message}");
                await LogAsync(correlationId, RecoverAction, "WARN", $"Could not send confirmation mail: {ex.Message}", request.TaskNumber, Severity.Warning, target);
            }

            await LogAsync(correlationId, RecoverAction, "DONE", "Recover Shared Mailbox has completed", request.TaskNumber, Severity.Success, target);
            Progress(onProgress, $"'{displayName}' restored.");

            return new RecoverSharedMailboxResult
            {
                Succeeded = true,
                DisplayName = displayName,
                PrimarySmtpAddress = address,
                AccessGroups = groupLines,
                CorrelationId = correlationId,
                WarningMessage = warnings.Count == 0 ? null : string.Join(" ", warnings),
            };
        }
        catch (Exception ex)
        {
            return await FailRecoverAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    // ----- gathering -----

    private async Task<RemovalState> GatherRemovalStateAsync(string mailboxIdentity, CancellationToken ct)
    {
        var identity = (mailboxIdentity ?? "").Trim();
        if (identity.Length == 0)
            throw new InvalidOperationException("Enter the shared mailbox's display name or address.");

        var (mailbox, resolveError) = await ResolveMailboxAsync(identity, ct);
        if (mailbox is null)
            throw new InvalidOperationException(resolveError);

        var displayName = Str(mailbox, "DisplayName");

        // Get-Mailbox finds user and room mailboxes just as readily. This page must never be the
        // way someone's personal mailbox gets deleted by a typo.
        var type = Str(mailbox, "RecipientTypeDetails");
        if (!string.Equals(type, "SharedMailbox", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException($"'{displayName}' is a {type}, not a shared mailbox. Only shared mailboxes can be removed here.");

        var address = Str(mailbox, "PrimarySmtpAddress");
        var atIndex = address.IndexOf('@');
        var domain = atIndex > 0 ? address[(atIndex + 1)..] : "";

        if (!Guid.TryParse(Str(mailbox, "ExchangeGuid"), out var exchangeGuid))
            throw new InvalidOperationException($"'{displayName}' has no ExchangeGuid - it cannot be removed safely from here.");

        var groups = new List<SharedMailboxAccessGroupInfo>();
        var discovered = await DiscoverAccessGroupsAsync(address, displayName, domain, ct);
        foreach (var tier in Tiers)
        {
            if (!discovered.TryGetValue(tier, out var groupName))
                continue;

            var groupAddress = (await PsValues.ExpandProperty(_host, "Get-DistributionGroup", groupName, "PrimarySmtpAddress", ct)).FirstOrDefault() ?? "";
            var members = await _host.InvokeAsync(ps => ps
                .AddCommand("Get-DistributionGroupMember")
                .AddParameter("Identity", groupName)
                .AddParameter("ResultSize", "Unlimited")
                .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

            groups.Add(new SharedMailboxAccessGroupInfo
            {
                Tier = tier,
                Name = groupName,
                Address = groupAddress,
                MemberNames = members.Select(m => Or(Str(m, "DisplayName"), Str(m, "Name"))).ToList(),
                MemberIdentities = members.Select(m => Or(Str(m, "PrimarySmtpAddress"), Str(m, "Name"))).Where(v => v.Length > 0).ToList(),
            });
        }

        // The .ED group's ManagedBy is the authoritative owner list, as everywhere else here; a
        // mailbox that only has .AU/.RE falls back to whichever group it does have.
        var ownerGroup = groups.FirstOrDefault(g => g.Tier == "ED") ?? groups.FirstOrDefault();
        var owners = ownerGroup is null
            ? new List<(string Upn, string DisplayName, string Mail)>()
            : await ResolveOwnerIdentitiesAsync(
                await PsValues.ExpandProperty(_host, "Get-DistributionGroup", ownerGroup.Name, "ManagedBy", ct), ct);

        // ForwardingSmtpAddress reads "smtp:x@y"; ForwardingAddress a recipient name. Either one
        // means mail stops going somewhere when this mailbox disappears.
        var forwarding = Str(mailbox, "ForwardingSmtpAddress");
        if (forwarding.StartsWith("smtp:", StringComparison.OrdinalIgnoreCase))
            forwarding = forwarding[5..];
        if (forwarding.Length == 0)
            forwarding = Str(mailbox, "ForwardingAddress");

        return new RemovalState(mailbox, displayName, address, exchangeGuid, owners, forwarding, groups,
            await GetDirectAccessAsync(address, ct));
    }

    /// <summary>
    /// Folder permissions held by someone other than an access group - the legacy dialog's
    /// "NonStandard Access". Default/Anonymous and "None" are not access.
    /// </summary>
    private async Task<List<string>> GetDirectAccessAsync(string address, CancellationToken ct)
    {
        var holders = new List<string>();
        foreach (var folder in new[] { $"{address}:\\", $"{address}:\\Calendar" })
        {
            IReadOnlyList<PSObject> perms;
            try
            {
                perms = await _host.InvokeAsync(ps => ps
                    .AddCommand("Get-MailboxFolderPermission")
                    .AddParameter("Identity", folder)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
            }
            catch
            {
                continue;
            }

            foreach (var perm in perms)
            {
                var user = UserValue(perm, "User");
                var rights = Value(perm, "AccessRights");
                if (user.Length == 0
                    || user.Equals("Default", StringComparison.OrdinalIgnoreCase)
                    || user.Equals("Anonymous", StringComparison.OrdinalIgnoreCase)
                    || rights.Equals("None", StringComparison.OrdinalIgnoreCase)
                    || user.StartsWith("MBX.", StringComparison.OrdinalIgnoreCase)
                    || Tiers.Any(t => user.EndsWith($".{t}", StringComparison.OrdinalIgnoreCase)))
                    continue;

                var line = $"{user} ({rights})";
                if (!holders.Contains(line, StringComparer.OrdinalIgnoreCase))
                    holders.Add(line);
            }
        }
        return holders;
    }

    /// <summary>
    /// The legacy report, section by section (ShrMbxRemove.ps1 wrote the same cmdlets' output to a
    /// text file), plus each access group with its owners and members - those are deleted for good.
    /// Every section is best-effort: an unreadable one is noted, not fatal.
    /// </summary>
    private async Task<string> BuildRemovalSnapshotAsync(RemovalState state, CancellationToken ct)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Pre-removal snapshot for {state.DisplayName} <{state.Address}> ({DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC)");

        AppendObjects(sb, "Mailbox", new[] { state.Mailbox }, new[]
        {
            "DisplayName", "Alias", "PrimarySmtpAddress", "EmailAddresses", "ExchangeGuid", "RecipientTypeDetails",
            "ForwardingAddress", "ForwardingSmtpAddress", "DeliverToMailboxAndForward", "GrantSendOnBehalfTo",
            "RequireSenderAuthenticationEnabled", "MailTip", "RetentionPolicy", "WhenCreated",
        });

        sb.AppendLine().AppendLine("--- Owners ---");
        sb.AppendLine(state.Owners.Count == 0 ? "  (none resolved)" : "  " + string.Join(", ", state.Owners.Select(o => $"{o.DisplayName} <{o.Upn}>")));

        foreach (var group in state.Groups)
        {
            sb.AppendLine().AppendLine($"--- Access group .{group.Tier}: {group.Name} <{group.Address}> ---");
            try
            {
                var groupObject = await _host.InvokeAsync(ps => ps
                    .AddCommand("Get-Group")
                    .AddParameter("Identity", group.Name)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                foreach (var g in groupObject)
                {
                    sb.AppendLine($"  ManagedBy: {Value(g, "ManagedBy")}");
                    sb.AppendLine($"  Notes: {Value(g, "Notes")}");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"  (could not read group details: {ex.Message})");
            }
            sb.AppendLine($"  Members ({group.MemberNames.Count}):");
            for (var i = 0; i < group.MemberNames.Count; i++)
            {
                var id = i < group.MemberIdentities.Count ? group.MemberIdentities[i] : "";
                sb.AppendLine($"    {group.MemberNames[i]} <{id}>");
            }
        }

        sb.AppendLine().AppendLine("--- Direct folder access (not through an access group) ---");
        sb.AppendLine(state.DirectAccess.Count == 0 ? "  (none)" : "  " + string.Join(Environment.NewLine + "  ", state.DirectAccess));

        await AppendCommandAsync(sb, "Mailbox permissions", "Get-MailboxPermission", state.Address, new[] { "User", "AccessRights", "IsInherited" },
            o => !Value(o, "User").StartsWith("NT AUTHORITY", StringComparison.OrdinalIgnoreCase)
                 && !string.Equals(Value(o, "IsInherited"), "True", StringComparison.OrdinalIgnoreCase), ct);
        await AppendCommandAsync(sb, "Recipient permissions (Send As)", "Get-RecipientPermission", state.Address, new[] { "Trustee", "AccessRights" },
            o => !Value(o, "Trustee").StartsWith("NT AUTHORITY", StringComparison.OrdinalIgnoreCase), ct);
        await AppendCommandAsync(sb, "Mailbox statistics", "Get-MailboxStatistics", state.Address,
            new[] { "ItemCount", "TotalItemSize", "LastLogonTime", "LastUserActionTime" }, null, ct);
        await AppendCommandAsync(sb, "Mailbox folder statistics", "Get-MailboxFolderStatistics", state.Address,
            new[] { "FolderPath", "ItemsInFolder" }, null, ct);

        return sb.ToString().TrimEnd();
    }

    private async Task AppendCommandAsync(StringBuilder sb, string title, string command, string identity, string[] properties,
        Func<PSObject, bool>? filter, CancellationToken ct)
    {
        try
        {
            var results = await _host.InvokeAsync(ps => ps
                .AddCommand(command)
                .AddParameter("Identity", identity)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            AppendObjects(sb, title, filter is null ? results : results.Where(filter), properties);
        }
        catch (Exception ex)
        {
            sb.AppendLine().AppendLine($"--- {title} ---").AppendLine($"  (could not read: {ex.Message})");
        }
    }

    private static void AppendObjects(StringBuilder sb, string title, IEnumerable<PSObject> items, string[] properties)
    {
        sb.AppendLine().AppendLine($"--- {title} ---");
        var any = false;
        foreach (var item in items)
        {
            if (any)
                sb.AppendLine();
            any = true;
            foreach (var name in properties)
            {
                var value = name == "User" ? UserValue(item, name) : Value(item, name);
                if (value.Length > 0)
                    sb.AppendLine($"  {name}: {value}");
            }
        }
        if (!any)
            sb.AppendLine("  (none)");
    }

    /// <summary>A property as text, with collections joined - ToString() on one gives only its type name.</summary>
    private static string Value(PSObject o, string name)
    {
        var value = o.Properties[name]?.Value;
        if (value is PSObject wrapped)
            value = wrapped.BaseObject;

        return value switch
        {
            null => "",
            string s => s,
            IEnumerable items => string.Join(", ", items.Cast<object?>().Where(x => x is not null).Select(x => x!.ToString())),
            _ => value.ToString() ?? "",
        };
    }

    private static DateTime? DateValue(PSObject o, string name)
    {
        var value = o.Properties[name]?.Value;
        if (value is PSObject wrapped)
            value = wrapped.BaseObject;
        return value switch
        {
            DateTime d => d,
            _ when DateTime.TryParse(value?.ToString(), out var parsed) => parsed,
            _ => null,
        };
    }

    // ----- mail, cache, failure -----

    /// <summary>Follows the legacy SharedMailboxRemoval.oft template.</summary>
    private async Task SendRemovalConfirmationMailAsync(IReadOnlyList<string> ownerMails, string displayName, bool groupsRemoved,
        string taskNumber, CancellationToken ct)
    {
        var name = WebUtility_HtmlEscape(displayName);

        // The template always said "and the groups created to manage access"; that is only true
        // when they actually went.
        var what = groupsRemoved
            ? $"the shared mailbox <b>{name}</b> and the groups created to manage access to this mailbox have"
            : $"the shared mailbox <b>{name}</b> has";

        var html = $"""
            <p>As requested {what} been removed from the system. It may take up to 72 hrs for the removal of
            these items to sync to the Offline Address List. In the meanwhile if individuals try to email this
            mailbox they will receive a delivery failure message indicating that the email address cannot be
            resolved.</p>

            <p>If the mailbox is needed after all, please contact the Service Desk within 30 days &ndash; until
            then it can be restored.</p>

            <p>With the removal of these items this ticket is being closed.</p>
            """;

        await SendMailToOwnersAsync(ownerMails, $"Removal of {displayName} Shared Mailbox - {taskNumber}", html, ct);
    }

    /// <summary>Flags the row rather than deleting it: the grid's "Deleted" column shows it, and a recovery clears it again.</summary>
    private async Task MarkCacheDeletedAsync(Guid exchangeGuid, CancellationToken ct)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return;

        try
        {
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync(ct);
            await using var cmd = new SqlCommand(
                "UPDATE dbo.SharedMailboxes SET IsDeletedInM365 = 1, LastSeenAtUtc = SYSUTCDATETIME() WHERE ExchangeGuid = @Id", conn);
            cmd.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = exchangeGuid;
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch
        {
            // Cache write is best-effort; the scheduled import will reconcile this row regardless.
        }
    }

    private async Task<RemoveSharedMailboxResult> FailRemoveAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, RemoveAction, "FAIL", message, taskNumber, Severity.Error);
        return RemoveSharedMailboxResult.Failed(correlationId, message);
    }

    private async Task<RecoverSharedMailboxResult> FailRecoverAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, RecoverAction, "FAIL", message, taskNumber, Severity.Error);
        return RecoverSharedMailboxResult.Failed(correlationId, message);
    }
}
