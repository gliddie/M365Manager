using System.Collections;
using System.Data;
using System.Management.Automation;
using System.Text.Json;
using M365Manager.Core.M365;
using M365Manager.Core.Naming;
using M365Manager.Core.Notifications;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Data.Logging;
using Microsoft.Data.SqlClient;

namespace M365Manager.Core.SharedMailboxes;

/// <summary>
/// See <see cref="ISharedMailboxService"/>. Owner resolution/Graph REST calls follow the same
/// pattern as <see cref="M365Manager.Core.Teams.TeamsService"/> (plain REST via
/// <see cref="IM365AuthService.GetAccessTokenAsync"/> for endpoints - here, /domains - not covered
/// well by the Kiota SDK version in use). Mailbox/group manipulation goes through
/// <see cref="PowerShellHost"/> against the same bundled Exchange Online module <see cref="M365Manager.Core.Exchange.ExchangeService"/> connects.
/// </summary>
public sealed class SharedMailboxService : ISharedMailboxService
{
    /// <summary>
    /// System folders that get no permission, matching the legacy Add-FolderPermissions loop.
    /// "Top of Information Store" is deliberately NOT here: legacy grants it too, addressing it as
    /// the mailbox root ("&lt;address&gt;:\"), which is the only folder access the .AU tier gets.
    /// </summary>
    private static readonly string[] SkipFolders = { "Recoverable Items", "Calendar Logging", "Deletions", "Purges", "Versions" };

    private readonly PowerShellHost _host;
    private readonly IM365AuthService _auth;
    private readonly ISettingsService _settings;
    private readonly ISharedMailboxNamingService _naming;
    private readonly ILogService _log;
    private readonly IConnectionStringProvider _connectionStrings;
    private readonly INotificationMailService _mail;
    private readonly GraphRestClient _graph;

    public SharedMailboxService(
        PowerShellHost host,
        IM365AuthService auth,
        ISettingsService settings,
        ISharedMailboxNamingService naming,
        ILogService log,
        IConnectionStringProvider connectionStrings,
        INotificationMailService mail,
        GraphRestClient graph)
    {
        _graph = graph;
        _host = host;
        _auth = auth;
        _settings = settings;
        _naming = naming;
        _log = log;
        _connectionStrings = connectionStrings;
        _mail = mail;
    }

    public async Task<SharedMailboxCreationResult> CreateSharedMailboxAsync(SharedMailboxCreationRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "CreateSharedMailbox", "STAR", "New Shared Mailbox creation has started", request.TaskNumber, Severity.Info);
            Progress(onProgress, "Starting shared mailbox creation...");

            var owners = await ResolveOwnersAsync(request.OwnerIdentities, ct);
            if (owners.Count == 0)
                return await FailAsync(correlationId, request.TaskNumber, $"Could not resolve any owner from '{request.OwnerIdentities}' to a directory user.");

            var ownerUpns = owners.Select(o => o.Upn).ToArray();
            var ownerMails = owners.Select(o => MailOrUpn(o.Upn, o.Mail)).ToArray();
            var ownerMailTip = BuildOwnerMailTip(owners.Select(o => o.DisplayName).ToList());
            Progress(onProgress, $"Owner(s) resolved: {string.Join(", ", ownerUpns)}");

            // The convention supplies the defaults; anything the operator corrected in the form
            // wins. The alias is always re-derived from the final address and sanitized, since
            // Exchange rejects an alias containing "&", umlauts and the like.
            var (computedDisplayName, computedLocalPart) = await _naming.BuildNameAsync(request.Location, request.Name, ct);

            var displayName = string.IsNullOrWhiteSpace(request.DisplayNameOverride)
                ? computedDisplayName
                : request.DisplayNameOverride.Trim();

            var address = string.IsNullOrWhiteSpace(request.AddressOverride)
                ? _naming.BuildAddress(computedLocalPart, request.Domain)
                : request.AddressOverride.Trim();

            var atIndex = address.IndexOf('@');
            if (atIndex <= 0)
                return await FailAsync(correlationId, request.TaskNumber, $"'{address}' is not a valid e-mail address.");
            if (displayName.Length == 0)
                return await FailAsync(correlationId, request.TaskNumber, "The mailbox needs a display name.");

            var alias = ExchangeNaming.ToAliasSafe(address[..atIndex]);
            Progress(onProgress, $"Mailbox name: {displayName} ({address})");

            var existing = await _host.InvokeAsync(ps => ps
                .AddCommand("Get-Mailbox")
                .AddParameter("Identity", address)
                .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);
            if (existing.Count > 0)
                return await FailAsync(correlationId, request.TaskNumber, $"A mailbox already exists for '{address}'.");

            await LogAsync(correlationId, "CreateSharedMailbox", "INPUT",
                $"Name={displayName}; Address={address}; Owners={string.Join(", ", ownerUpns)}; AllowExternalSenders={request.AllowExternalSenders}",
                request.TaskNumber, Severity.Info);

            Progress(onProgress, "Creating shared mailbox...");
            var created = await _host.InvokeAsync(ps => ps
                .AddCommand("New-Mailbox")
                .AddParameter("Name", displayName)
                .AddParameter("Shared", true)
                .AddParameter("Alias", alias)
                .AddParameter("PrimarySmtpAddress", address)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            await LogAsync(correlationId, "CreateSharedMailbox", "CREATE", $"{displayName} ({address})", request.TaskNumber, Severity.Success);

            Progress(onProgress, "Disabling IMAP/POP...");
            await DisableProtocolsAsync(address, ct);
            await LogAsync(correlationId, "CreateSharedMailbox", "DISAB", "Disabled IMAP/POP", request.TaskNumber, Severity.Success);

            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-Mailbox")
                .AddParameter("Identity", address)
                .AddParameter("MessageCopyForSendOnBehalfEnabled", true)
                .AddParameter("MessageCopyForSentAsEnabled", true)
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            var smSettings = _settings.Current.SharedMailboxes;
            if (!string.IsNullOrWhiteSpace(smSettings.RetentionPolicyName) || !string.IsNullOrWhiteSpace(smSettings.RoleAssignmentPolicyName))
            {
                Progress(onProgress, "Applying configured retention/role assignment policy...");
                await _host.InvokeAsync(ps =>
                {
                    ps.AddCommand("Set-Mailbox").AddParameter("Identity", address);
                    if (!string.IsNullOrWhiteSpace(smSettings.RetentionPolicyName))
                    {
                        ps.AddParameter("RetentionPolicy", smSettings.RetentionPolicyName);
                        ps.AddParameter("RetentionHoldEnabled", false);
                    }
                    if (!string.IsNullOrWhiteSpace(smSettings.RoleAssignmentPolicyName))
                        ps.AddParameter("RoleAssignmentPolicy", smSettings.RoleAssignmentPolicyName);
                    ps.AddParameter("ErrorAction", "Stop");
                }, ct: ct);
                await LogAsync(correlationId, "CreateSharedMailbox", "CHNG", "Applied configured retention/role assignment policy", request.TaskNumber, Severity.Success);
            }

            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-Mailbox")
                .AddParameter("Identity", address)
                .AddParameter("RequireSenderAuthenticationEnabled", !request.AllowExternalSenders)
                .AddParameter("MailTip", ownerMailTip)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            await LogAsync(correlationId, "CreateSharedMailbox", "CHNG",
                $"External senders {(request.AllowExternalSenders ? "allowed" : "blocked")}", request.TaskNumber, Severity.Success);

            // The Editor group is always created - it carries FullAccess/SendAs and is what
            // "owner" resolves to later. If no explicit .ED members were given, the owners
            // themselves go in, so the mailbox is never left without anyone who can open it.
            // The Author and Reader tiers are only created when someone was actually named for
            // them; empty ones would just be clutter.
            var editorMembers = string.IsNullOrWhiteSpace(request.EditorMembers)
                ? request.OwnerIdentities
                : request.EditorMembers;

            var tiers = new List<(string Tier, string FolderRight, string MembersCsv)>
            {
                ("ED", "Editor", editorMembers),
            };
            if (!string.IsNullOrWhiteSpace(request.AuthorMembers))
                tiers.Add(("AU", "PublishingAuthor", request.AuthorMembers));
            if (!string.IsNullOrWhiteSpace(request.ReaderMembers))
                tiers.Add(("RE", "Reviewer", request.ReaderMembers));

            string? edGroupAddress = null, auGroupAddress = null, reGroupAddress = null;
            var createdTiers = new List<AccessTier>();

            foreach (var (tier, folderRight, membersCsv) in tiers)
            {
                Progress(onProgress, $"Creating .{tier} access group...");

                // Name/display name keep the mailbox display name's spaces; the alias and address
                // are the same string with the spaces removed.
                var groupName = _naming.BuildAccessGroupName(displayName, tier);
                var groupAlias = _naming.BuildAccessGroupAlias(displayName, tier);
                var groupAddress = _naming.BuildAccessGroupAddress(displayName, tier, request.Domain);

                await _host.InvokeAsync(ps => ps
                    .AddCommand("New-DistributionGroup")
                    .AddParameter("Name", groupName)
                    .AddParameter("DisplayName", groupName)
                    .AddParameter("PrimarySmtpAddress", groupAddress)
                    .AddParameter("Alias", groupAlias)
                    .AddParameter("ManagedBy", ownerUpns)
                    .AddParameter("Type", "Security")
                    .AddParameter("RequireSenderAuthenticationEnabled", true)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);

                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-Group")
                    .AddParameter("Identity", groupName)
                    .AddParameter("Notes", $"{ownerMailTip} - Per: {request.TaskNumber}")
                    .AddParameter("ErrorAction", "Stop"), ct: ct);

                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-DistributionGroup")
                    .AddParameter("Identity", groupName)
                    .AddParameter("MailTip", ownerMailTip)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);

                await LogAsync(correlationId, "CreateSharedMailbox", "GRP", $"Created {groupName} ({groupAddress})", request.TaskNumber, Severity.Success);

                Progress(onProgress, $"Granting .{tier} mailbox permissions...");
                await GrantTierMailboxRightsAsync(address, groupAddress, tier, ct);
                await GrantFolderPermissionsAsync(address, groupAddress, folderRight, ct);
                await LogAsync(correlationId, "CreateSharedMailbox", "PERM", $"Granted {folderRight} rights to {groupAddress}", request.TaskNumber, Severity.Success);

                if (!string.IsNullOrWhiteSpace(membersCsv))
                {
                    await AddGroupMembersAsync(groupAddress, membersCsv, ct);
                    await LogAsync(correlationId, "CreateSharedMailbox", "MEMBER", $"Added members to {groupAddress}: {membersCsv}", request.TaskNumber, Severity.Success);
                }

                createdTiers.Add(new AccessTier(tier, groupName, membersCsv));

                switch (tier)
                {
                    case "ED": edGroupAddress = groupAddress; break;
                    case "AU": auGroupAddress = groupAddress; break;
                    case "RE": reGroupAddress = groupAddress; break;
                }
            }

            // Best-effort: a mailbox that exists but whose owners weren't told is still a success.
            // The failure is reported all the way back to the caller though - it used to be logged
            // only, so a relay/mailbox problem went unnoticed for as long as nobody read the log.
            string? mailWarning = null;
            Progress(onProgress, $"Sending confirmation e-mail to the owner(s) as {_mail.SenderDescription}...");
            try
            {
                await SendCreationConfirmationMailAsync(owners, displayName, address, request.TaskNumber,
                    request.AllowExternalSenders, createdTiers, ct);
                await LogAsync(correlationId, "CreateSharedMailbox", "SEND",
                    $"Sent confirmation mail via {_mail.SenderDescription} to {string.Join(", ", ownerMails)}",
                    request.TaskNumber, Severity.Success);
            }
            catch (Exception ex)
            {
                mailWarning = $"The mailbox was created, but the confirmation e-mail to {string.Join(", ", ownerMails)} could not be sent: {ex.Message}";
                await LogAsync(correlationId, "CreateSharedMailbox", "WARN", $"Could not send confirmation mail: {ex.Message}", request.TaskNumber, Severity.Warning);
                Progress(onProgress, $"WARNING: {mailWarning}");
            }

            await LogAsync(correlationId, "CreateSharedMailbox", "DONE", "New Shared Mailbox creation has completed", request.TaskNumber, Severity.Success);
            Progress(onProgress, $"Shared mailbox '{displayName}' created successfully.");

            await UpsertCacheAsync(created.Count > 0 ? created[0] : null, displayName, alias, address,
                edGroupAddress, auGroupAddress, reGroupAddress, string.Join("; ", ownerUpns),
                !request.AllowExternalSenders, ct);

            return new SharedMailboxCreationResult
            {
                Succeeded = true,
                DisplayName = displayName,
                PrimarySmtpAddress = address,
                CorrelationId = correlationId,
                WarningMessage = mailWarning,
            };
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    public async Task<ChangeSharedMailboxOwnerResult> ChangeOwnerAsync(ChangeSharedMailboxOwnerRequest request, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "ChangeOwner", "STAR", "Change Shared Mailbox Owner has started", request.TaskNumber, Severity.Info);

            var address = (request.MailboxAddress ?? "").Trim();
            var atIndex = address.IndexOf('@');
            if (atIndex <= 0)
                return await FailChangeAsync(correlationId, request.TaskNumber, $"'{request.MailboxAddress}' is not a valid mailbox address.");

            var domain = address[(atIndex + 1)..];

            // The access groups are named after the mailbox's DISPLAY name ("MBX.<display name>.ED"),
            // not its address, so the display name has to come from Exchange - it can't be derived
            // from the address, whose local part has the spaces stripped and a dot inserted.
            var mailbox = await _host.InvokeAsync(ps => ps
                .AddCommand("Get-Mailbox")
                .AddParameter("Identity", address)
                .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);
            if (mailbox.Count == 0)
                return await FailChangeAsync(correlationId, request.TaskNumber, $"No mailbox found for '{address}'.");

            var mailboxDisplayName = Str(mailbox[0], "DisplayName");
            if (mailboxDisplayName.Length == 0)
                return await FailChangeAsync(correlationId, request.TaskNumber, $"Mailbox '{address}' has no display name to derive its access group names from.");

            var newOwnerUpns = SplitIdentities(request.OwnerIdentities).Select(AppendDomain).ToArray();
            if (newOwnerUpns.Length == 0)
                return await FailChangeAsync(correlationId, request.TaskNumber, "Enter at least one owner identity.");

            await LogAsync(correlationId, "ChangeOwner", "INPUT",
                $"Mailbox={address}; Mode={request.Mode}; Owners={string.Join(", ", newOwnerUpns)}", request.TaskNumber, Severity.Info);

            var accessGroups = await DiscoverAccessGroupsAsync(address, mailboxDisplayName, domain, ct);
            if (accessGroups.Count == 0)
            {
                return await FailChangeAsync(correlationId, request.TaskNumber,
                    $"No .ED/.AU/.RE access group found for '{address}'. Neither the mailbox's permissions nor the names " +
                    $"the convention produces ('{_naming.BuildAccessGroupName(mailboxDisplayName, "ED")}' / " +
                    $"'{_naming.BuildAccessGroupAddress(mailboxDisplayName, "ED", domain)}') matched an existing group.");
            }

            await LogAsync(correlationId, "ChangeOwner", "FOUND",
                "Access groups: " + string.Join(", ", accessGroups.Select(g => $"{g.Key}={g.Value}")),
                request.TaskNumber, Severity.Info);

            var touchedGroups = new List<string>();
            var finalDisplayNames = new List<string>();
            var finalOwnerUpns = new List<string>();
            // Kept apart from the UPNs on purpose: the UPNs go into the cache, these into the To: line.
            var finalOwnerMails = new List<string>();

            foreach (var tier in Tiers)
            {
                if (!accessGroups.TryGetValue(tier, out var groupName))
                    continue;

                await ApplyOwnerChangeAsync(groupName, request.Mode, newOwnerUpns, ct);
                await LogAsync(correlationId, "ChangeOwner", "CHGOWN",
                    $"{request.Mode} owner(s) {string.Join(", ", newOwnerUpns)} on {groupName}", request.TaskNumber, Severity.Success);

                // Expanded rather than read off the group object - each value is used as an
                // identity below, and the property has been seen arriving flattened into one string.
                var currentManagedBy = await PsValues.ExpandProperty(_host, "Get-DistributionGroup", groupName, "ManagedBy", ct);
                var owners = await ResolveOwnerIdentitiesAsync(currentManagedBy, ct);
                var displayNames = owners.Select(o => o.DisplayName).ToList();
                var mailTip = BuildOwnerMailTip(displayNames);

                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-DistributionGroup")
                    .AddParameter("Identity", groupName)
                    .AddParameter("MailTip", mailTip)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-Group")
                    .AddParameter("Identity", groupName)
                    .AddParameter("Notes", $"{mailTip} - Per: {request.TaskNumber}")
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                await LogAsync(correlationId, "ChangeOwner", "UPGTIP", $"Updated {groupName} MailTip/Notes to '{mailTip}'", request.TaskNumber, Severity.Success);

                touchedGroups.Add(groupName);

                // The .ED group is the authoritative owner list - it is what the cache's Owners
                // column and the mailbox MailTip document. Taking whichever tier happened to be
                // processed last would let .RE overwrite it.
                if (tier == "ED" || finalOwnerUpns.Count == 0)
                {
                    finalDisplayNames = displayNames;
                    finalOwnerUpns = owners.Select(o => o.Upn).ToList();
                    finalOwnerMails = owners.Select(o => MailOrUpn(o.Upn, o.Mail)).ToList();
                }
            }

            var mailboxMailTip = BuildOwnerMailTip(finalDisplayNames);
            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-Mailbox")
                .AddParameter("Identity", address)
                .AddParameter("MailTip", mailboxMailTip)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            await LogAsync(correlationId, "ChangeOwner", "UPMTIP", $"Updated mailbox MailTip to '{mailboxMailTip}'", request.TaskNumber, Severity.Success);

            // UPNs, not display names - CreateSharedMailboxAsync fills this column with UPNs and the
            // import script writes ManagedBy addresses, so the grid must not switch format here.
            await UpsertOwnerCacheAsync(address, string.Join("; ", finalOwnerUpns), ct);

            string? mailWarning = null;
            try
            {
                // Only the tiers that actually exist on this mailbox, in ED/AU/RE order.
                var existingTiers = Tiers
                    .Where(accessGroups.ContainsKey)
                    .Select(tier => new AccessTier(tier, accessGroups[tier], ""))
                    .ToArray();

                await SendOwnerChangeConfirmationMailAsync(finalOwnerMails, finalDisplayNames,
                    mailboxDisplayName, address, existingTiers, request.TaskNumber, ct);
                await LogAsync(correlationId, "ChangeOwner", "SEND",
                    $"Sent confirmation mail via {_mail.SenderDescription} to {string.Join(", ", finalOwnerMails)}",
                    request.TaskNumber, Severity.Success);
            }
            catch (Exception ex)
            {
                mailWarning = $"The owners were changed, but the confirmation e-mail to {string.Join(", ", finalOwnerMails)} could not be sent: {ex.Message}";
                await LogAsync(correlationId, "ChangeOwner", "WARN", $"Could not send confirmation mail: {ex.Message}", request.TaskNumber, Severity.Warning);
            }

            await LogAsync(correlationId, "ChangeOwner", "DONE", "Change Shared Mailbox Owner has completed", request.TaskNumber, Severity.Success);

            return new ChangeSharedMailboxOwnerResult
            {
                Succeeded = true,
                UpdatedOwners = mailboxMailTip,
                CorrelationId = correlationId,
                WarningMessage = mailWarning,
            };
        }
        catch (Exception ex)
        {
            return await FailChangeAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    private static readonly string[] Tiers = { "ED", "AU", "RE" };

    /// <summary>
    /// Finds a mailbox's .ED/.AU/.RE access groups, returning tier -> group identity.
    ///
    /// The groups are looked up from the permissions actually granted on the mailbox, the way
    /// ShrMbxChgOwner.ps1 did, rather than from the names the current convention would produce.
    /// Deriving the name breaks on any mailbox whose groups were named under an older convention
    /// or renamed by hand afterwards: renaming a group's display name in the admin center does NOT
    /// change its SMTP address, so "MBX.UL Arkansas Communication.ED" can still answer to
    /// "MBX.UL.ArkansasCommunication.ED@ul.com" and a lookup by the newly derived address misses it.
    ///
    /// The derived names are still tried afterwards, for groups that exist but hold no permission
    /// on the mailbox yet.
    /// </summary>
    private async Task<Dictionary<string, string>> DiscoverAccessGroupsAsync(
        string address, string mailboxDisplayName, string domain, CancellationToken ct)
    {
        var found = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // 1. Top-of-store permissions - where the .ED group's FullAccess lives.
        try
        {
            var perms = await _host.InvokeAsync(ps => ps
                .AddCommand("Get-MailboxPermission")
                .AddParameter("Identity", address)
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            foreach (var perm in perms)
                RecordTier(found, UserValue(perm, "User"));
        }
        catch
        {
            // fall through to the other lookups
        }

        // 2. Folder-level permissions - the only place .AU and .RE ever appear.
        foreach (var folder in new[] { $"{address}:\\", $"{address}:\\Calendar" })
        {
            try
            {
                var perms = await _host.InvokeAsync(ps => ps
                    .AddCommand("Get-MailboxFolderPermission")
                    .AddParameter("Identity", folder)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);

                foreach (var perm in perms)
                    RecordTier(found, UserValue(perm, "User"));
            }
            catch
            {
                // folder may not exist / not be readable - try the next one
            }
        }

        // 3. Anything still missing: fall back to the names the convention produces, by address
        //    first and then by name, since a hand-renamed group answers to one but not the other.
        foreach (var tier in Tiers)
        {
            if (found.ContainsKey(tier))
                continue;

            foreach (var candidate in new[]
                     {
                         _naming.BuildAccessGroupAddress(mailboxDisplayName, tier, domain),
                         _naming.BuildAccessGroupName(mailboxDisplayName, tier),
                     })
            {
                if (await ResolveGroupAsync(candidate, ct) is { } identity)
                {
                    found[tier] = identity;
                    break;
                }
            }
        }

        // Everything discovered from permissions still has to be a real distribution group -
        // a stray FullAccess for a user account called "...ED" must not be treated as one.
        var verified = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (tier, candidate) in found)
        {
            if (await ResolveGroupAsync(candidate, ct) is { } identity)
                verified[tier] = identity;
        }

        return verified;
    }

    /// <summary>Returns the group's canonical Name when the identity resolves to a distribution group, else null.</summary>
    private async Task<string?> ResolveGroupAsync(string identity, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(identity))
            return null;

        var result = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-DistributionGroup")
            .AddParameter("Identity", identity)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

        if (result.Count == 0)
            return null;

        var name = Str(result[0], "Name");
        return name.Length > 0 ? name : identity;
    }

    /// <summary>Records a permission holder whose name ends in ".ED"/".AU"/".RE" under that tier.</summary>
    private static void RecordTier(IDictionary<string, string> found, string rawUser)
    {
        if (string.IsNullOrWhiteSpace(rawUser))
            return;

        // Get-MailboxPermission reports users as "domain/OU/Name" - keep the leaf.
        var leaf = rawUser.Split('/').Last().Trim();
        if (leaf.Length == 0)
            return;

        foreach (var tier in Tiers)
        {
            if (leaf.EndsWith($".{tier}", StringComparison.OrdinalIgnoreCase) && !found.ContainsKey(tier))
                found[tier] = leaf;
        }
    }

    /// <summary>
    /// Get-MailboxPermission returns User as a plain string, Get-MailboxFolderPermission as an
    /// object whose DisplayName carries the group name - handle both.
    /// </summary>
    private static string UserValue(PSObject o, string name)
    {
        var value = o.Properties[name]?.Value;
        if (value is null)
            return "";

        if (value is PSObject nested)
        {
            var displayName = nested.Properties["DisplayName"]?.Value?.ToString();
            if (!string.IsNullOrWhiteSpace(displayName))
                return displayName;
        }

        return value.ToString() ?? "";
    }

    private async Task ApplyOwnerChangeAsync(string groupName, OwnerChangeMode mode, string[] owners, CancellationToken ct)
    {
        switch (mode)
        {
            case OwnerChangeMode.Replace:
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-DistributionGroup")
                    .AddParameter("Identity", groupName)
                    .AddParameter("ManagedBy", owners)
                    .AddParameter("BypassSecurityGroupManagerCheck", true)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                break;
            case OwnerChangeMode.Add:
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-DistributionGroup")
                    .AddParameter("Identity", groupName)
                    .AddParameter("ManagedBy", new Hashtable { ["add"] = owners })
                    .AddParameter("BypassSecurityGroupManagerCheck", true)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                break;
            case OwnerChangeMode.Remove:
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-DistributionGroup")
                    .AddParameter("Identity", groupName)
                    .AddParameter("ManagedBy", new Hashtable { ["remove"] = owners })
                    .AddParameter("BypassSecurityGroupManagerCheck", true)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                break;
        }
    }

    public async Task<IReadOnlyList<string>> GetAvailableDomainsAsync(CancellationToken ct = default)
    {
        using var json = await _graph.GetAsync("/domains", ct: ct);

        var domains = json.RootElement.GetProperty("value").EnumerateArray()
            .Where(d => d.TryGetProperty("isVerified", out var v) && v.GetBoolean())
            .Where(d => d.TryGetProperty("supportedServices", out var s) &&
                        s.EnumerateArray().Any(x => string.Equals(x.GetString(), "Email", StringComparison.OrdinalIgnoreCase)))
            .Select(d => d.GetProperty("id").GetString() ?? "")
            .Where(id => id.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(d => d, StringComparer.OrdinalIgnoreCase)
            .ToList();

        // DefaultMailDomain, not M365.Domain (the UPN/sign-in domain) - see AppendDomain vs. this.
        var defaultDomain = _settings.Current.M365.DefaultMailDomain.Trim();
        if (defaultDomain.Length > 0 && domains.RemoveAll(d => string.Equals(d, defaultDomain, StringComparison.OrdinalIgnoreCase)) > 0)
            domains.Insert(0, defaultDomain);

        return domains;
    }

    // ----- steps -----

    private async Task DisableProtocolsAsync(string identity, CancellationToken ct)
    {
        for (var attempt = 1; attempt <= 5; attempt++)
        {
            try
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-CASMailbox")
                    .AddParameter("Identity", identity)
                    .AddParameter("ImapEnabled", false)
                    .AddParameter("PopEnabled", false)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                return;
            }
            catch (Exception) when (attempt < 5)
            {
                await Task.Delay(TimeSpan.FromSeconds(3), ct);
            }
        }
    }

    private async Task GrantTierMailboxRightsAsync(string mailboxAddress, string groupAddress, string tier, CancellationToken ct)
    {
        if (tier == "RE")
            return; // Reader gets folder-level Reviewer only - no top-of-store access.

        // Additive update syntax (same technique as the ManagedBy add/remove calls in
        // ApplyOwnerChangeAsync) instead of read-current-then-append-then-overwrite: on a
        // freshly created mailbox, GrantSendOnBehalfTo's "unset" value round-trips through this
        // module as something that fails RecipientIdParameter conversion once reassigned as a
        // plain array, even though it reads back as an empty collection. Add-only sidesteps
        // reading the existing value entirely.
        await _host.InvokeAsync(ps => ps
            .AddCommand("Set-Mailbox")
            .AddParameter("Identity", mailboxAddress)
            .AddParameter("GrantSendOnBehalfTo", new Hashtable { ["Add"] = groupAddress })
            .AddParameter("ErrorAction", "Stop"), ct: ct);

        if (tier != "ED")
            return;

        for (var attempt = 1; attempt <= 5; attempt++)
        {
            try
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Add-MailboxPermission")
                    .AddParameter("Identity", mailboxAddress)
                    .AddParameter("User", groupAddress)
                    .AddParameter("AccessRights", "FullAccess")
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                break;
            }
            catch (Exception) when (attempt < 5)
            {
                await Task.Delay(TimeSpan.FromSeconds(3), ct);
            }
        }

        await _host.InvokeAsync(ps => ps
            .AddCommand("Add-RecipientPermission")
            .AddParameter("Identity", mailboxAddress)
            .AddParameter("Trustee", groupAddress)
            .AddParameter("AccessRights", "SendAs")
            .AddParameter("Confirm", false)
            .AddParameter("ErrorAction", "Stop"), ct: ct);
    }

    /// <summary>
    /// Exchange folder permissions don't inherit to subfolders, so - like the legacy
    /// ShrMbxNew.ps1 Add-FolderPermissions loop - every folder is granted individually.
    /// Best-effort per folder: a locked/system folder shouldn't abort the whole grant.
    /// </summary>
    private async Task GrantFolderPermissionsAsync(string mailboxAddress, string groupAddress, string accessRight, CancellationToken ct)
    {
        IReadOnlyList<PSObject> folders;
        try
        {
            folders = await _host.InvokeAsync(ps => ps
                .AddCommand("Get-MailboxFolderStatistics")
                .AddParameter("Identity", mailboxAddress)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
        }
        catch
        {
            return; // Top-of-store permission (ED/AU) still grants access even if this fails.
        }

        foreach (var folder in folders)
        {
            var identityRaw = Str(folder, "Identity");
            var parts = identityRaw.Split('\\');
            if (parts.Length < 2)
                continue;

            var relative = string.Join("\\", parts.Skip(1));
            var leaf = parts[^1];
            if (SkipFolders.Contains(leaf, StringComparer.OrdinalIgnoreCase))
                continue;

            var folderIdentity = relative.Equals("Top of Information Store", StringComparison.OrdinalIgnoreCase)
                ? $"{mailboxAddress}:\\"
                : $"{mailboxAddress}:\\{relative}";

            try
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Add-MailboxFolderPermission")
                    .AddParameter("Identity", folderIdentity)
                    .AddParameter("User", groupAddress)
                    .AddParameter("AccessRights", accessRight)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
            }
            catch
            {
                // best-effort per folder
            }
        }
    }

    private async Task AddGroupMembersAsync(string groupAddress, string membersCsv, CancellationToken ct)
    {
        foreach (var raw in SplitIdentities(membersCsv))
        {
            var resolved = AppendDomain(raw);
            try
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Add-DistributionGroupMember")
                    .AddParameter("Identity", groupAddress)
                    .AddParameter("Member", resolved)
                    .AddParameter("BypassSecurityGroupManagerCheck", true)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
            }
            catch
            {
                // best-effort per member - a single bad identity shouldn't abort the whole request
            }
        }
    }

    // ----- owner resolution -----

    private async Task<List<(string Upn, string DisplayName, string Mail)>> ResolveOwnersAsync(string ownersCsv, CancellationToken ct)
    {
        var resolved = new List<(string Upn, string DisplayName, string Mail)>();
        foreach (var raw in SplitIdentities(ownersCsv))
        {
            var owner = await ResolveUserAsync(raw, ct);
            if (owner is not null)
                resolved.Add((owner.Value.Upn, owner.Value.DisplayName, owner.Value.Mail));
        }
        return resolved;
    }

    /// <summary>
    /// The address a notification should actually be sent to: the directory's mail attribute, or
    /// the UPN when the directory has none. Sending to a UPN only ever worked because Graph
    /// /me/sendMail resolves it inside the tenant - an external SMTP relay cannot.
    /// </summary>
    private static string MailOrUpn(string upn, string mail) =>
        string.IsNullOrWhiteSpace(mail) ? upn : mail;

    /// <summary>
    /// Resolves whatever Exchange reports in ManagedBy - which is a recipient Name/DN, NOT an SMTP
    /// address - to the real UPN and display name. Sending mail to, or caching, the raw ManagedBy
    /// value would fail (Graph rejects it as an address) or write a different identifier format
    /// than the creation path stores.
    /// </summary>
    private async Task<List<(string Upn, string DisplayName, string Mail)>> ResolveOwnerIdentitiesAsync(IReadOnlyList<string> identities, CancellationToken ct)
    {
        var resolved = new List<(string Upn, string DisplayName, string Mail)>();
        foreach (var id in identities)
        {
            var user = await ResolveUserAsync(id, ct);
            if (user is not null)
            {
                resolved.Add((user.Value.Upn, user.Value.DisplayName, user.Value.Mail));
                continue;
            }

            // Graph couldn't resolve it (ManagedBy often carries the Exchange display name rather
            // than anything Graph indexes) - ask Exchange itself for the SMTP address.
            var recipient = await _host.InvokeAsync(ps => ps
                .AddCommand("Get-Recipient")
                .AddParameter("Identity", id)
                .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

            if (recipient.Count > 0)
            {
                // Exchange's PrimarySmtpAddress is already a deliverable address, so it serves as
                // both the identifier and the mail target here.
                var smtp = Str(recipient[0], "PrimarySmtpAddress");
                var display = Str(recipient[0], "DisplayName");
                resolved.Add((smtp.Length > 0 ? smtp : id, display.Length > 0 ? display : id, smtp));
            }
            else
            {
                resolved.Add((id, id, ""));
            }
        }
        return resolved;
    }

    /// <summary>
    /// The UPN and the mail address are different things in this tenant - UPNs sit on the sign-in
    /// domain (e.g. 34632@global.ul.com), mailboxes on the mail domain (e.g. first.last@ul.com).
    /// Both are returned: the UPN identifies the user everywhere this code stores or compares
    /// owners, <see cref="Mail"/> is the only one an SMTP relay can actually deliver to.
    /// <see cref="Mail"/> is empty when the directory has none.
    /// </summary>
    private async Task<(string Id, string Upn, string DisplayName, string Mail)?> ResolveUserAsync(string identity, CancellationToken ct)
    {
        var value = (identity ?? "").Trim();
        if (value.Length == 0)
            return null;

        var candidate = AppendDomain(value);

        // Via GraphRestClient rather than the Kiota SDK client, so the lookup shows up in the
        // console next to the Exchange cmdlets instead of happening invisibly.
        const string select = "$select=id,userPrincipalName,displayName,mail";

        try
        {
            using var json = await _graph.GetAsync($"/users/{Uri.EscapeDataString(candidate)}?{select}", ct: ct);
            var user = json.RootElement;
            var id = GraphText(user, "id");
            if (id.Length > 0)
                return (id,
                    Or(GraphText(user, "userPrincipalName"), candidate),
                    Or(GraphText(user, "displayName"), candidate),
                    GraphText(user, "mail"));
        }
        catch
        {
            // fall through to mail-filter lookup below
        }

        if (value.Contains('@'))
        {
            var filter = Uri.EscapeDataString($"mail eq '{value.Replace("'", "''")}'");
            using var json = await _graph.GetAsync($"/users?$filter={filter}&{select}", ct: ct);

            if (json.RootElement.TryGetProperty("value", out var matches) && matches.GetArrayLength() > 0)
            {
                var match = matches[0];
                var id = GraphText(match, "id");
                if (id.Length > 0)
                    return (id,
                        Or(GraphText(match, "userPrincipalName"), value),
                        Or(GraphText(match, "displayName"), value),
                        Or(GraphText(match, "mail"), value));
            }
        }

        return null;
    }

    private static string GraphText(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : "";

    private static string Or(string value, string fallback) => value.Length > 0 ? value : fallback;

    /// <summary>A bare SamAccountName (no "@") is completed with the configured default domain.</summary>
    private string AppendDomain(string identity)
    {
        var value = (identity ?? "").Trim();
        if (value.Length == 0 || value.Contains('@'))
            return value;

        var domain = _settings.Current.M365.Domain.Trim();
        return domain.Length == 0 ? value : $"{value}@{domain}";
    }

    private static IEnumerable<string> SplitIdentities(string csv)
        => (csv ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    // Mirrors the legacy "Owner: First Last, ..." mail tip, truncated to Exchange's 175-char limit.
    private static string BuildOwnerMailTip(IReadOnlyList<string> displayNames)
    {
        var text = "Owner: " + string.Join(", ", displayNames);
        return text.Length > 175 ? text[..175] : text;
    }

    // ----- owner notification e-mail -----
    //
    // The wording follows the legacy Outlook templates the service desk used to send by hand
    // (SharedMailboxNew.oft / SharedMailboxOwnershipChange.oft, extracted to docs/mail-templates/
    // by scripts/Export-MailTemplates.ps1). Their manual placeholders - "NameOfMailbox",
    // "MBX.???.ED", "NamesofIndividualsGivenEditorAccess" - are filled from the request here.
    // The screenshot SharedMailboxNew.oft embedded is deliberately not carried over: it showed the
    // Outlook 2010 address book, including the sending admin's own mail address.

    /// <summary>One access tier as it should appear in the confirmation mail.</summary>
    private readonly record struct AccessTier(string Tier, string GroupName, string MembersCsv);

    private static readonly Dictionary<string, string> TierDescriptions = new()
    {
        ["ED"] = "grants full access to the mailbox",
        ["AU"] = "grants access to create new folders, read and respond to all messages but can delete and/or manage only those messages they are the author of",
        ["RE"] = "grants read only access to the mailbox",
    };

    private static readonly Dictionary<string, string> TierAccessLabels = new()
    {
        ["ED"] = "Editor",
        ["AU"] = "Author",
        ["RE"] = "Read",
    };

    private async Task SendCreationConfirmationMailAsync(
        IReadOnlyList<(string Upn, string DisplayName, string Mail)> owners,
        string displayName,
        string address,
        string taskNumber,
        bool allowExternalSenders,
        IReadOnlyList<AccessTier> tiers,
        CancellationToken ct)
    {
        var settings = _settings.Current.SharedMailboxes;
        var name = WebUtility_HtmlEscape(displayName);

        var retention = string.IsNullOrWhiteSpace(settings.RetentionPolicyName)
            ? "has been created and configured"
            : $"has been created and configured with the retention policy \"{WebUtility_HtmlEscape(settings.RetentionPolicyName)}\"";

        var external = allowExternalSenders
            ? "This mailbox accepts e-mail from external senders."
            : "Since your request indicated that this mailbox address would not be shared with clients, "
              + "this mailbox has been configured to not allow receipt of e-mail from external sources.";

        var groupList = string.Concat(tiers.Select(t =>
            $"<li><b>{WebUtility_HtmlEscape(t.GroupName)}</b> &ndash; {TierDescriptions[t.Tier]}</li>"));

        // The members are whatever the operator typed - usually SamAccountNames. The recipient of
        // this mail needs names, not employee numbers.
        var accessListBlocks = new List<string>();
        foreach (var tier in tiers.Where(t => !string.IsNullOrWhiteSpace(t.MembersCsv)))
        {
            var names = await RecipientLookup.ToDisplayNamesAsync(_host, SplitIdentities(tier.MembersCsv), ct);
            accessListBlocks.Add($"<p>Individuals granted {TierAccessLabels[tier.Tier]} Access:<br>"
                                 + $"{WebUtility_HtmlEscape(string.Join(", ", names))}</p>");
        }
        var accessLists = string.Concat(accessListBlocks);

        var html = $"""
            <p>Per the subject service desk ticket the shared mailbox <b>{name}</b> {retention}.
            {external}
            It may take up to 72 hours for the address for the new mailbox and group(s) that grant(s) access
            to sync to the Offline Address List. In the meanwhile you can find this mailbox by selecting the
            Global Address List.</p>

            <p>As requested the following individuals have been made the mailbox owner:<br>
            {WebUtility_HtmlEscape(string.Join(", ", owners.Select(o => o.DisplayName)))}</p>

            <p>As the owner of the mailbox you have been granted the rights to update the membership of the
            group(s) that control(s) access to the mailbox. Mailbox ownership does not give you rights to
            access the mailbox. If you did not list yourself in one of the access groups in your request and
            you need access to the mailbox you will need to add yourself as a member of one of the group(s)
            shown below. By default we only create groups where individuals have been identified for access.
            An individual should only be added to the group which grants the highest level of access.</p>

            <p>To grant others access to the mailbox you will modify the membership of the group(s):</p>
            <ul>{groupList}</ul>
            {BuildGroupMembershipLink(settings)}

            <p><b>Note</b> &ndash; Shared mailboxes that have not been accessed for 6 months (180 days) by an
            individual that has been granted access will automatically be purged from the system.</p>

            <p>Please use this information to add this mailbox to your Outlook Navigator. Additionally, if you
            grant others access to this mailbox please forward the below information to them so they can add
            the mailbox to their Outlook Navigator.</p>
            {accessLists}
            <p>Mailbox Name: <b>{name}</b><br>
            Email Address: <b>{WebUtility_HtmlEscape(address)}</b></p>
            {BuildHelpLinks(settings)}

            <p>In the future if this mailbox is no longer needed/used please create a request with the Service
            Desk so the mailbox can be removed from the system.</p>

            <p>With this information this ticket is being closed.</p>
            """;

        await SendMailToOwnersAsync(
            owners.Select(o => MailOrUpn(o.Upn, o.Mail)).ToArray(),
            $"{displayName} Shared Mailbox - {taskNumber}",
            html, ct);
    }

    private async Task SendOwnerChangeConfirmationMailAsync(
        IReadOnlyList<string> ownerUpns,
        IReadOnlyList<string> ownerDisplayNames,
        string displayName,
        string address,
        IReadOnlyList<AccessTier> tiers,
        string taskNumber,
        CancellationToken ct)
    {
        if (ownerUpns.Count == 0)
            return;

        var settings = _settings.Current.SharedMailboxes;
        var name = WebUtility_HtmlEscape(displayName);

        // The legacy template says "MBX.???.ED is the group you will edit to give others ...",
        // with the operator replacing the ??? by hand. The discovered group names go in instead.
        var groupList = string.Concat(tiers.Select(t =>
            $"<li><b>{WebUtility_HtmlEscape(t.GroupName)}</b> &ndash; {TierDescriptions[t.Tier]}</li>"));

        var html = $"""
            <p>Per the subject service desk ticket the ownership of the <b>{name}</b> shared mailbox has been
            changed. You have been made an/the owner of the mailbox.</p>

            <p>As owner of the mailbox you are responsible for updating the membership of the group that
            controls access to the mailbox. Mailbox ownership does not give you rights to access the mailbox.
            If you need access to the mailbox you will need to add yourself as a member of one of the group(s)
            shown below. An individual should only be added to the group which grants the highest level of
            access. To modify access you will update the membership of:</p>
            <ul>{groupList}</ul>

            <p>By default we only create groups where individuals have been identified for access. If you need
            to have a group created that grants lesser access such as Author or Reader please create a Service
            Desk request so that we can create the groups necessary for those access levels.</p>
            {BuildGroupMembershipLink(settings)}

            <p>Current owner(s): {WebUtility_HtmlEscape(string.Join(", ", ownerDisplayNames))}</p>

            <p><b>Note</b> &ndash; Shared mailboxes that have not been accessed for 6 months (180 days) by an
            individual that has been granted access will automatically be purged from the system.</p>

            <p>Please use this information to add this mailbox to your Outlook Navigator. Additionally, if you
            grant others access to this mailbox please forward the below information to them so they can add
            the mailbox to their Outlook Navigator.</p>

            <p>Mailbox Name: <b>{name}</b><br>
            Email Address: <b>{WebUtility_HtmlEscape(address)}</b></p>
            {BuildHelpLinks(settings)}

            <p>With this information this ticket is being closed.</p>
            """;

        await SendMailToOwnersAsync(
            ownerUpns,
            $"{displayName} Shared Mailbox Ownership Change - {taskNumber}",
            html, ct);
    }

    private static string BuildGroupMembershipLink(SharedMailboxSettings settings) =>
        string.IsNullOrWhiteSpace(settings.GroupMembershipHelpUrl)
            ? ""
            : $"""<p><a href="{WebUtility_HtmlEscape(settings.GroupMembershipHelpUrl.Trim())}">Editing Group Membership</a></p>""";

    /// <summary>Renders the "Label|URL" lines from settings as a list of links; empty when unset.</summary>
    private static string BuildHelpLinks(SharedMailboxSettings settings)
    {
        var links = (settings.MailboxHelpLinks ?? "")
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(line => line.Split('|', 2, StringSplitOptions.TrimEntries))
            .Where(parts => parts.Length == 2 && parts[0].Length > 0 && parts[1].Length > 0)
            .Select(parts => $"""<li><a href="{WebUtility_HtmlEscape(parts[1])}">{WebUtility_HtmlEscape(parts[0])}</a></li>""")
            .ToArray();

        return links.Length == 0 ? "" : $"<ul>{string.Concat(links)}</ul>";
    }

    private Task SendMailToOwnersAsync(IReadOnlyList<string> ownerUpns, string subject, string htmlBody, CancellationToken ct)
        => _mail.SendAsync(ownerUpns, subject, htmlBody, ct: ct);

    private static string WebUtility_HtmlEscape(string value) => System.Net.WebUtility.HtmlEncode(value);

    // ----- SQL cache (read for the overview grid, write-through after creation/owner change) -----

    public async Task<IReadOnlyList<SharedMailboxOverviewRow>> GetSharedMailboxesOverviewAsync(CancellationToken ct = default)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return Array.Empty<SharedMailboxOverviewRow>();

        const string query = """
            SELECT
                ExchangeGuid, DisplayName, Alias, PrimarySmtpAddress, Owners,
                EDGroupAddress, AUGroupAddress, REGroupAddress,
                RequireSenderAuthenticationEnabled, IsDeletedInM365, CreatedDateTime, LastImportedAtUtc
            FROM dbo.SharedMailboxes
            ORDER BY DisplayName;
            """;

        await using var conn = new SqlConnection(cs);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(query, conn);
        await using var reader = await cmd.ExecuteReaderAsync(ct);

        var list = new List<SharedMailboxOverviewRow>();
        while (await reader.ReadAsync(ct))
        {
            list.Add(new SharedMailboxOverviewRow
            {
                ExchangeGuid = reader.GetGuid(0).ToString(),
                DisplayName = reader.IsDBNull(1) ? "" : reader.GetString(1),
                Alias = reader.IsDBNull(2) ? "" : reader.GetString(2),
                PrimarySmtpAddress = reader.IsDBNull(3) ? "" : reader.GetString(3),
                Owners = reader.IsDBNull(4) ? "" : reader.GetString(4).Replace(';', ',').Replace(",", ", "),
                EDGroupAddress = reader.IsDBNull(5) ? "" : reader.GetString(5),
                AUGroupAddress = reader.IsDBNull(6) ? "" : reader.GetString(6),
                REGroupAddress = reader.IsDBNull(7) ? "" : reader.GetString(7),
                RequireSenderAuthenticationEnabled = reader.IsDBNull(8) ? null : reader.GetBoolean(8),
                IsDeletedInM365 = !reader.IsDBNull(9) && reader.GetBoolean(9),
                CreatedDateTime = reader.IsDBNull(10) ? null : reader.GetDateTime(10),
                LastImportedAtUtc = reader.IsDBNull(11) ? default : reader.GetDateTime(11),
            });
        }
        return list;
    }

    private async Task UpsertCacheAsync(PSObject? mailbox, string displayName, string alias, string address,
        string? edGroupAddress, string? auGroupAddress, string? reGroupAddress, string owners,
        bool requireSenderAuthenticationEnabled, CancellationToken ct)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return;

        var guidStr = mailbox is not null ? Str(mailbox, "ExchangeGuid") : "";
        if (!Guid.TryParse(guidStr, out var exchangeGuid))
            return;

        try
        {
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync(ct);

            const string sql = """
                MERGE dbo.SharedMailboxes AS target
                USING (SELECT @Id AS ExchangeGuid) AS source ON target.ExchangeGuid = source.ExchangeGuid
                WHEN MATCHED THEN UPDATE SET
                    DisplayName = @DisplayName, Alias = @Alias, PrimarySmtpAddress = @PrimarySmtpAddress,
                    EDGroupAddress = @EDGroupAddress, AUGroupAddress = @AUGroupAddress, REGroupAddress = @REGroupAddress,
                    Owners = @Owners, RequireSenderAuthenticationEnabled = @RequireSenderAuthenticationEnabled,
                    IsDeletedInM365 = 0, LastImportedAtUtc = SYSUTCDATETIME(), LastSeenAtUtc = SYSUTCDATETIME()
                WHEN NOT MATCHED THEN INSERT
                    (ExchangeGuid, DisplayName, Alias, PrimarySmtpAddress, EDGroupAddress, AUGroupAddress, REGroupAddress,
                     Owners, RequireSenderAuthenticationEnabled, CreatedDateTime, IsDeletedInM365, LastImportedAtUtc, LastSeenAtUtc)
                VALUES
                    (@Id, @DisplayName, @Alias, @PrimarySmtpAddress, @EDGroupAddress, @AUGroupAddress, @REGroupAddress,
                     @Owners, @RequireSenderAuthenticationEnabled, SYSUTCDATETIME(), 0, SYSUTCDATETIME(), SYSUTCDATETIME());
                """;

            await using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = exchangeGuid;
            cmd.Parameters.Add("@DisplayName", SqlDbType.NVarChar, 256).Value = displayName;
            cmd.Parameters.Add("@Alias", SqlDbType.NVarChar, 128).Value = alias;
            cmd.Parameters.Add("@PrimarySmtpAddress", SqlDbType.NVarChar, 256).Value = address;
            cmd.Parameters.Add("@EDGroupAddress", SqlDbType.NVarChar, 256).Value = NullableParam(edGroupAddress ?? "");
            cmd.Parameters.Add("@AUGroupAddress", SqlDbType.NVarChar, 256).Value = NullableParam(auGroupAddress ?? "");
            cmd.Parameters.Add("@REGroupAddress", SqlDbType.NVarChar, 256).Value = NullableParam(reGroupAddress ?? "");
            cmd.Parameters.Add("@Owners", SqlDbType.NVarChar, -1).Value = owners;
            cmd.Parameters.Add("@RequireSenderAuthenticationEnabled", SqlDbType.Bit).Value = requireSenderAuthenticationEnabled;
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch
        {
            // Cache write is best-effort; the scheduled import will reconcile this row regardless.
        }
    }

    private async Task UpsertOwnerCacheAsync(string address, string owners, CancellationToken ct)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return;

        try
        {
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync(ct);
            await using var cmd = new SqlCommand(
                "UPDATE dbo.SharedMailboxes SET Owners = @Owners, LastSeenAtUtc = SYSUTCDATETIME() WHERE PrimarySmtpAddress = @Address", conn);
            cmd.Parameters.Add("@Owners", SqlDbType.NVarChar, -1).Value = owners;
            cmd.Parameters.Add("@Address", SqlDbType.NVarChar, 256).Value = address;
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch
        {
            // Cache write is best-effort; the scheduled import will reconcile this row regardless.
        }
    }

    private static object NullableParam(string value) => value.Length > 0 ? value : DBNull.Value;

    // ----- logging -----

    private async Task<SharedMailboxCreationResult> FailAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, "CreateSharedMailbox", "FAIL", message, taskNumber, Severity.Error);
        return SharedMailboxCreationResult.Failed(correlationId, message);
    }

    private async Task<ChangeSharedMailboxOwnerResult> FailChangeAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, "ChangeOwner", "FAIL", message, taskNumber, Severity.Error);
        return ChangeSharedMailboxOwnerResult.Failed(correlationId, message);
    }

    private async Task LogAsync(Guid correlationId, string action, string eventCode, string message, string taskNumber, Severity severity)
    {
        try
        {
            await _log.WriteAsync(new LogEntry
            {
                Area = "SharedMailboxes",
                Action = action,
                EventCode = eventCode,
                Severity = severity,
                Message = message,
                TaskNumber = taskNumber,
                CorrelationId = correlationId,
                UserUpn = _auth.CurrentUser?.Upn,
            });
        }
        catch
        {
            // Logging is best-effort; never let it break the creation/change flow.
        }
    }

    private static void Progress(Action<string>? onProgress, string message) => onProgress?.Invoke(message);

    private static string Str(PSObject? o, string name) => o?.Properties[name]?.Value?.ToString() ?? "";

    private static List<string> StrList(PSObject o, string name) => PsValues.ToStringList(o, name);
}
