using System.Collections;
using System.Data;
using System.Management.Automation;
using System.Text;
using M365Manager.Core.Exchange;
using M365Manager.Core.M365;
using M365Manager.Core.Naming;
using M365Manager.Core.Notifications;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Data.Logging;
using Microsoft.Data.SqlClient;

namespace M365Manager.Core.Groups;

/// <summary>
/// See <see cref="IGroupAdminService"/>. Everything runs through <see cref="PowerShellHost"/> on
/// the Exchange Online connection <see cref="ExchangeService"/> already established - no Graph, no
/// second sign-in.
/// </summary>
public sealed class GroupAdminService : IGroupAdminService
{
    /// <summary>Exchange truncates a MailTip beyond this, so do it ourselves and keep it readable.</summary>
    private const int MailTipLimit = 175;

    /// <summary>Exchange's limit on a group Name/Alias.</summary>
    private const int NameLimit = 64;

    private readonly PowerShellHost _host;
    private readonly IM365AuthService _auth;
    private readonly ISettingsService _settings;
    private readonly IGroupNamingService _naming;
    private readonly ILogService _log;
    private readonly IConnectionStringProvider _connectionStrings;
    private readonly INotificationMailService _mail;

    public GroupAdminService(
        PowerShellHost host,
        IM365AuthService auth,
        ISettingsService settings,
        IGroupNamingService naming,
        ILogService log,
        IConnectionStringProvider connectionStrings,
        INotificationMailService mail)
    {
        _mail = mail;
        _host = host;
        _auth = auth;
        _settings = settings;
        _naming = naming;
        _log = log;
        _connectionStrings = connectionStrings;
    }

    // ----- preview -----

    public async Task<GroupCreationPreview> PreviewAsync(GroupCreationRequest request, CancellationToken ct = default)
    {
        var names = await _naming.BuildAsync(request.RawName, request.IsTemporary, request.RemoveCharacter, request.Domain, ct);
        var errors = new List<string>();

        var displayName = string.IsNullOrWhiteSpace(request.DisplayNameOverride) ? names.DisplayName : request.DisplayNameOverride.Trim();
        var address = string.IsNullOrWhiteSpace(request.AddressOverride) ? names.Address : request.AddressOverride.Trim();

        if (displayName.Length == 0)
            errors.Add("Enter the group name in dot notation, e.g. \"LST.EMEA.Some Team\".");
        else if (displayName.Length > NameLimit)
            errors.Add($"The group name is {displayName.Length} characters - Exchange allows at most {NameLimit}.");

        if (address.Length > 0 && !address.Contains('@'))
            errors.Add($"'{address}' is not a valid e-mail address.");

        if (string.IsNullOrWhiteSpace(request.Owners))
            errors.Add("Enter at least one owner.");

        var exists = address.Length > 0 ? await TryProbeAsync(address, ct) : null;
        if (exists == true)
            errors.Add($"A group already exists at {address}.");

        return new GroupCreationPreview
        {
            Names = new GroupNames(displayName, names.Alias, address),
            GroupExists = exists,
            Errors = errors,
        };
    }

    /// <summary>Distinguishes "does not exist" (false) from "could not tell" (null) - see RoomResourceService.</summary>
    private async Task<bool?> TryProbeAsync(string identity, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(identity))
            return null;

        try
        {
            return await GetRecipientAsync(identity, ct) is not null;
        }
        catch
        {
            return null;
        }
    }

    // ----- create -----

    public async Task<GroupOperationResult> CreateAsync(GroupCreationRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "Create", "STAR", "New group creation has started", request.TaskNumber, Severity.Info);
            Progress(onProgress, "Validating...");

            var preview = await PreviewAsync(request, ct);
            if (!preview.CanCreate)
                return await FailAsync(correlationId, "Create", request.TaskNumber, string.Join(" ", preview.Errors));

            if (preview.HasInconclusiveProbes)
                return await FailAsync(correlationId, "Create", request.TaskNumber,
                    "Could not verify in Exchange Online whether this group already exists. Check the connection and try again - nothing was created.");

            var displayName = preview.Names.DisplayName;
            var address = preview.Names.Address;
            var alias = ExchangeNaming.ToAliasSafe(address[..address.IndexOf('@')]);
            var owners = SplitIdentities(request.Owners).Select(AppendDomain).ToArray();

            await LogAsync(correlationId, "Create", "INPUT",
                $"Name={displayName}; Address={address}; Kind={request.Kind}; Owners={string.Join(", ", owners)}; " +
                $"Temporary={request.IsTemporary}; AllowExternalSenders={request.AllowExternalSenders}",
                request.TaskNumber, Severity.Info);

            // 1. The group itself.
            Progress(onProgress, $"Creating {request.Kind.ToString().ToLowerInvariant()} group {displayName}...");
            await _host.InvokeAsync(ps => ps
                .AddCommand("New-DistributionGroup")
                .AddParameter("Name", displayName)
                .AddParameter("PrimarySmtpAddress", address)
                .AddParameter("Alias", alias)
                .AddParameter("ManagedBy", owners)
                .AddParameter("Type", request.Kind.ToString())
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            await LogAsync(correlationId, "Create", "CREATE", $"{displayName} ({address})", request.TaskNumber, Severity.Success);

            // 2. MailTip and Notes from the resolved owners.
            var ownerMailTip = await BuildOwnerMailTipAsync(displayName, correlationId, "Create", request.TaskNumber, ct);
            var notes = request.IsTemporary && request.RemoveOn is { } removeOn
                ? $"{ownerMailTip} - Per: {request.TaskNumber} - Remove On: {removeOn:MM/dd/yyyy}"
                : $"{ownerMailTip} - Per: {request.TaskNumber}";

            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-DistributionGroup")
                .AddParameter("Identity", displayName)
                .AddParameter("MailTip", ownerMailTip)
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-Group")
                .AddParameter("Identity", displayName)
                .AddParameter("Notes", notes)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            await LogAsync(correlationId, "Create", "CFG", $"MailTip and Notes set: {notes}", request.TaskNumber, Severity.Success);

            // 3. Sender authentication, provenance stamp and any extra aliases.
            await _host.InvokeAsync(ps =>
            {
                ps.AddCommand("Set-DistributionGroup")
                  .AddParameter("Identity", address)
                  .AddParameter("RequireSenderAuthenticationEnabled", !request.AllowExternalSenders)
                  .AddParameter("BypassSecurityGroupManagerCheck", true)
                  .AddParameter("CustomAttribute15", $"M365Manager NewDistributionGroup {DateTime.UtcNow:yyyy-MM-dd HH:mm} Per: {request.TaskNumber}")
                  .AddParameter("ErrorAction", "Stop");

                var aliases = SplitIdentities(request.AdditionalAliases).ToArray();
                if (aliases.Length > 0)
                    ps.AddParameter("EmailAddresses", new Hashtable { ["add"] = aliases });
            }, ct: ct);
            await LogAsync(correlationId, "Create", "CFG",
                $"External senders {(request.AllowExternalSenders ? "allowed" : "blocked")}", request.TaskNumber, Severity.Success);

            var results = new List<MemberOperationResult>();

            // 4. Authorized senders.
            if (!string.IsNullOrWhiteSpace(request.AuthorizedSenders))
            {
                Progress(onProgress, "Adding authorized senders...");
                var senderResults = await ApplyListAsync(address, "AcceptMessagesOnlyFromSendersOrMembers",
                    ListChangeMode.Add, SplitIdentities(request.AuthorizedSenders), isDynamic: false, ct);
                results.AddRange(senderResults);
                await LogResultsAsync(correlationId, "Create", "AUTH", "authorized sender", senderResults, request.TaskNumber);
            }

            // 5. Members.
            if (!string.IsNullOrWhiteSpace(request.Members))
            {
                Progress(onProgress, "Adding members...");
                var memberResults = await AddMembersAsync(address, SplitIdentities(request.Members), ct);
                results.AddRange(memberResults);
                await LogResultsAsync(correlationId, "Create", "MEMBER", "member", memberResults, request.TaskNumber);
            }

            string? mailWarning = null;
            if (!string.IsNullOrWhiteSpace(request.RequesterIdentity))
            {
                Progress(onProgress, $"Sending confirmation e-mail as {_mail.SenderDescription}...");
                var ownerNames = await ResolveOwnerNamesAsync(owners, ct);
                mailWarning = await NotifyRequesterAsync(
                    correlationId, "Create", request.RequesterIdentity, request.TaskNumber,
                    GroupNotificationMail.Creation(displayName, address,
                        SplitIdentities(request.AdditionalAliases).ToArray(), ownerNames, request.AllowExternalSenders,
                        _settings.Current.SharedMailboxes, request.TaskNumber), ct);
                if (mailWarning is not null)
                    Progress(onProgress, $"WARNING: {mailWarning}");
            }

            await LogAsync(correlationId, "Create", "DONE", "New group creation has completed", request.TaskNumber, Severity.Success);
            Progress(onProgress, $"{displayName} created.");

            await UpsertCacheAsync(displayName, alias, address, request.Kind, owners, ct);

            var failed = results.Count(r => !r.Succeeded);
            var memberWarning = failed > 0 ? $"{failed} identity/identities could not be added - see the results list." : null;

            return new GroupOperationResult
            {
                Succeeded = true,
                DisplayName = displayName,
                Address = address,
                Results = results,
                CorrelationId = correlationId,
                Warning = string.Join(" ", new[] { memberWarning, mailWarning }.Where(w => w is not null)) is { Length: > 0 } w ? w : null,
            };
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, "Create", request.TaskNumber, ex.Message);
        }
    }

    public async Task<GroupOperationResult> CreateDynamicAsync(DynamicGroupCreationRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "CreateDynamic", "STAR", "New dynamic group creation has started", request.TaskNumber, Severity.Info);

            var displayName = (request.DisplayName ?? "").Trim();
            if (displayName.Length == 0)
                return await FailAsync(correlationId, "CreateDynamic", request.TaskNumber, "Enter a display name.");
            if (displayName.Length > NameLimit)
                return await FailAsync(correlationId, "CreateDynamic", request.TaskNumber,
                    $"The name is {displayName.Length} characters - Exchange allows at most {NameLimit}.");

            var alias = ExchangeNaming.ToAliasSafe(displayName);
            var address = string.IsNullOrWhiteSpace(request.AddressOverride)
                ? $"{alias}@{request.Domain.Trim()}"
                : request.AddressOverride.Trim();

            if (await TryProbeAsync(address, ct) is not false)
                return await FailAsync(correlationId, "CreateDynamic", request.TaskNumber,
                    $"A recipient already exists at {address}, or its existence could not be verified.");

            await LogAsync(correlationId, "CreateDynamic", "INPUT",
                $"Name={displayName}; Address={address}; IncludedRecipients={request.IncludedRecipients}; " +
                $"Company={request.ConditionalCompany}; CA1={request.ConditionalCustomAttribute1}; " +
                $"CA2={request.ConditionalCustomAttribute2}; CA3={request.ConditionalCustomAttribute3}; CA8={request.ConditionalCustomAttribute8}",
                request.TaskNumber, Severity.Info);

            Progress(onProgress, $"Creating dynamic group {displayName}...");
            await _host.InvokeAsync(ps =>
            {
                ps.AddCommand("New-DynamicDistributionGroup")
                  .AddParameter("Name", displayName)
                  .AddParameter("DisplayName", displayName)
                  .AddParameter("Alias", alias)
                  .AddParameter("PrimarySmtpAddress", address)
                  .AddParameter("ErrorAction", "Stop");

                if (!string.IsNullOrWhiteSpace(request.IncludedRecipients))
                    ps.AddParameter("IncludedRecipients", request.IncludedRecipients.Trim());

                // Only the conditions that were filled in are sent; an empty one would narrow the
                // filter to recipients whose attribute is literally blank.
                AddIfSet(ps, "ConditionalCompany", request.ConditionalCompany);
                AddIfSet(ps, "ConditionalCustomAttribute1", request.ConditionalCustomAttribute1);
                AddIfSet(ps, "ConditionalCustomAttribute2", request.ConditionalCustomAttribute2);
                AddIfSet(ps, "ConditionalCustomAttribute3", request.ConditionalCustomAttribute3);
                AddIfSet(ps, "ConditionalCustomAttribute8", request.ConditionalCustomAttribute8);
            }, ct: ct);
            await LogAsync(correlationId, "CreateDynamic", "CREATE", $"{displayName} ({address})", request.TaskNumber, Severity.Success);

            if (!string.IsNullOrWhiteSpace(request.MailTip))
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-DynamicDistributionGroup")
                    .AddParameter("Identity", address)
                    .AddParameter("MailTip", Truncate(request.MailTip.Trim(), MailTipLimit))
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
            }

            if (!string.IsNullOrWhiteSpace(request.Notes))
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-Group")
                    .AddParameter("Identity", displayName)
                    .AddParameter("Notes", $"{request.Notes.Trim()} - Per: {request.TaskNumber}")
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
            }

            var mailWarning = await NotifyRequesterAsync(
                correlationId, "CreateDynamic", request.RequesterIdentity, request.TaskNumber,
                GroupNotificationMail.DynamicCreation(displayName, address, request.TaskNumber), ct);
            if (mailWarning is not null)
                Progress(onProgress, $"WARNING: {mailWarning}");

            await LogAsync(correlationId, "CreateDynamic", "DONE", "New dynamic group creation has completed", request.TaskNumber, Severity.Success);
            Progress(onProgress, $"{displayName} created.");

            return new GroupOperationResult
            {
                Succeeded = true,
                DisplayName = displayName,
                Address = address,
                CorrelationId = correlationId,
                Warning = mailWarning,
            };
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, "CreateDynamic", request.TaskNumber, ex.Message);
        }
    }

    private static void AddIfSet(System.Management.Automation.PowerShell ps, string parameter, string value)
    {
        if (!string.IsNullOrWhiteSpace(value))
            ps.AddParameter(parameter, value.Trim());
    }

    // ----- owners / authorized senders / aliases -----

    public async Task<GroupOperationResult> ChangeOwnersAsync(GroupListChangeRequest request, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "ChangeOwners", "STAR", "Owner change has started", request.TaskNumber, Severity.Info);

            var identities = SplitIdentities(request.Identities).Select(AppendDomain).ToArray();
            if (identities.Length == 0)
                return await FailAsync(correlationId, "ChangeOwners", request.TaskNumber, "Enter at least one owner identity.");

            await LogAsync(correlationId, "ChangeOwners", "INPUT",
                $"Group={request.GroupIdentity}; Mode={request.Mode}; Owners={string.Join(", ", identities)}", request.TaskNumber, Severity.Info);

            var results = new List<MemberOperationResult>();

            if (request.IsUnifiedGroup)
            {
                // M365 groups keep owners as links, not as a ManagedBy list.
                if (request.Mode == ListChangeMode.Replace)
                    return await FailAsync(correlationId, "ChangeOwners", request.TaskNumber,
                        "Replace is not supported for Microsoft 365 groups - add the new owners, then remove the old ones.");

                var command = request.Mode == ListChangeMode.Add ? "Add-UnifiedGroupLinks" : "Remove-UnifiedGroupLinks";
                foreach (var identity in identities)
                {
                    try
                    {
                        await _host.InvokeAsync(ps => ps
                            .AddCommand(command)
                            .AddParameter("Identity", request.GroupIdentity)
                            .AddParameter("LinkType", "Owners")
                            .AddParameter("Links", identity)
                            .AddParameter("Confirm", false)
                            .AddParameter("ErrorAction", "Stop"), ct: ct);
                        results.Add(new MemberOperationResult(identity, true, null));
                    }
                    catch (Exception ex)
                    {
                        results.Add(new MemberOperationResult(identity, false, ex.Message));
                    }
                }
            }
            else
            {
                results.AddRange(await ApplyListAsync(request.GroupIdentity, "ManagedBy", request.Mode, identities, isDynamic: false, ct));

                // Keep the MailTip/Notes in step with the new owner list, as the legacy scripts did.
                try
                {
                    var mailTip = await BuildOwnerMailTipAsync(request.GroupIdentity, correlationId, "ChangeOwners", request.TaskNumber, ct);
                    await _host.InvokeAsync(ps => ps
                        .AddCommand("Set-DistributionGroup")
                        .AddParameter("Identity", request.GroupIdentity)
                        .AddParameter("MailTip", mailTip)
                        .AddParameter("BypassSecurityGroupManagerCheck", true)
                        .AddParameter("ErrorAction", "Stop"), ct: ct);
                    await LogAsync(correlationId, "ChangeOwners", "UPGTIP", $"MailTip updated to '{mailTip}'", request.TaskNumber, Severity.Success);
                }
                catch (Exception ex)
                {
                    await LogAsync(correlationId, "ChangeOwners", "WARN", $"Could not refresh the MailTip: {ex.Message}", request.TaskNumber, Severity.Warning);
                }
            }

            await LogResultsAsync(correlationId, "ChangeOwners", "CHGOWN", "owner", results, request.TaskNumber);

            // Only the owners that actually changed belong in the mail.
            var changed = results.Where(r => r.Succeeded).Select(r => r.Identity).ToArray();
            string? mailWarning = null;
            if (changed.Length > 0)
            {
                var names = await ResolveOwnerNamesAsync(changed, ct);
                var displayName = (await GetDetailsAsync(request.GroupIdentity, ct))?.DisplayName ?? request.GroupIdentity;
                mailWarning = await NotifyRequesterAsync(
                    correlationId, "ChangeOwners", request.RequesterIdentity, request.TaskNumber,
                    GroupNotificationMail.OwnershipChanged(displayName, request.Mode, names,
                        request.IsUnifiedGroup, _settings.Current.SharedMailboxes, request.TaskNumber), ct);
            }

            await LogAsync(correlationId, "ChangeOwners", "DONE", "Owner change has completed", request.TaskNumber, Severity.Success);

            return Summarize(correlationId, request.GroupIdentity, results, mailWarning);
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, "ChangeOwners", request.TaskNumber, ex.Message);
        }
    }

    public async Task<GroupOperationResult> ChangeAuthorizedSendersAsync(GroupListChangeRequest request, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "ChangeAuthSenders", "STAR", "Authorized sender change has started", request.TaskNumber, Severity.Info);

            var identities = SplitIdentities(request.Identities).Select(AppendDomain).ToArray();
            if (identities.Length == 0)
                return await FailAsync(correlationId, "ChangeAuthSenders", request.TaskNumber, "Enter at least one identity.");

            var details = await GetDetailsAsync(request.GroupIdentity, ct);
            var isDynamic = details?.IsDynamic ?? false;

            await LogAsync(correlationId, "ChangeAuthSenders", "INPUT",
                $"Group={request.GroupIdentity}; Mode={request.Mode}; Identities={string.Join(", ", identities)}", request.TaskNumber, Severity.Info);

            var results = await ApplyListAsync(request.GroupIdentity, "AcceptMessagesOnlyFromSendersOrMembers",
                request.Mode, identities, isDynamic, ct);

            await LogResultsAsync(correlationId, "ChangeAuthSenders", "AUTH", "authorized sender", results, request.TaskNumber);

            var changed = results.Where(r => r.Succeeded).Select(r => r.Identity).ToArray();
            string? mailWarning = null;
            if (changed.Length > 0)
            {
                var names = await ResolveOwnerNamesAsync(changed, ct);
                mailWarning = await NotifyRequesterAsync(
                    correlationId, "ChangeAuthSenders", request.RequesterIdentity, request.TaskNumber,
                    GroupNotificationMail.AuthorizedSendersChanged(
                        details?.DisplayName ?? request.GroupIdentity, request.Mode, names, request.TaskNumber), ct);
            }

            await LogAsync(correlationId, "ChangeAuthSenders", "DONE", "Authorized sender change has completed", request.TaskNumber, Severity.Success);

            return Summarize(correlationId, request.GroupIdentity, results, mailWarning);
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, "ChangeAuthSenders", request.TaskNumber, ex.Message);
        }
    }

    public async Task<GroupOperationResult> ChangeAliasesAsync(GroupListChangeRequest request, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "ChangeAliases", "STAR", "Alias change has started", request.TaskNumber, Severity.Info);

            if (request.Mode == ListChangeMode.Replace)
                return await FailAsync(correlationId, "ChangeAliases", request.TaskNumber,
                    "Replace is not supported for aliases - it would drop the primary address. Add and remove individually.");

            var identities = SplitIdentities(request.Identities).ToArray();
            if (identities.Length == 0)
                return await FailAsync(correlationId, "ChangeAliases", request.TaskNumber, "Enter at least one address.");

            await LogAsync(correlationId, "ChangeAliases", "INPUT",
                $"Group={request.GroupIdentity}; Mode={request.Mode}; Addresses={string.Join(", ", identities)}", request.TaskNumber, Severity.Info);

            var results = await ApplyListAsync(request.GroupIdentity, "EmailAddresses", request.Mode, identities, isDynamic: false, ct);

            await LogResultsAsync(correlationId, "ChangeAliases", "ALIAS", "alias", results, request.TaskNumber);
            await LogAsync(correlationId, "ChangeAliases", "DONE", "Alias change has completed", request.TaskNumber, Severity.Success);

            return Summarize(correlationId, request.GroupIdentity, results);
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, "ChangeAliases", request.TaskNumber, ex.Message);
        }
    }

    /// <summary>
    /// Applies an add/remove/replace to a multi-valued Exchange property. Add and Remove use the
    /// additive @{add=}/@{remove=} syntax one identity at a time, so a single bad entry doesn't take
    /// the rest down with it; Replace assigns the whole array in one call.
    /// </summary>
    private async Task<List<MemberOperationResult>> ApplyListAsync(
        string groupIdentity, string property, ListChangeMode mode,
        IEnumerable<string> identities, bool isDynamic, CancellationToken ct)
    {
        var command = isDynamic ? "Set-DynamicDistributionGroup" : "Set-DistributionGroup";
        var results = new List<MemberOperationResult>();
        var list = identities.ToArray();

        if (mode == ListChangeMode.Replace)
        {
            try
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand(command)
                    .AddParameter("Identity", groupIdentity)
                    .AddParameter(property, list)
                    .AddParameter("BypassSecurityGroupManagerCheck", true)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);

                foreach (var identity in list)
                    results.Add(new MemberOperationResult(identity, true, null));
            }
            catch (Exception ex)
            {
                foreach (var identity in list)
                    results.Add(new MemberOperationResult(identity, false, ex.Message));
            }
            return results;
        }

        var key = mode == ListChangeMode.Add ? "add" : "remove";
        foreach (var identity in list)
        {
            try
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand(command)
                    .AddParameter("Identity", groupIdentity)
                    .AddParameter(property, new Hashtable { [key] = identity })
                    .AddParameter("BypassSecurityGroupManagerCheck", true)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                results.Add(new MemberOperationResult(identity, true, null));
            }
            catch (Exception ex)
            {
                results.Add(new MemberOperationResult(identity, false, ex.Message));
            }
        }
        return results;
    }

    // ----- membership -----

    public async Task<GroupOperationResult> ReplaceMembersAsync(GroupListChangeRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "ReplaceMembers", "STAR", "Membership replacement has started", request.TaskNumber, Severity.Info);

            var identities = SplitIdentities(request.Identities).Select(AppendDomain).ToArray();

            // Update-DistributionGroupMember with an empty array empties the group. Emptying a
            // list is a legitimate thing to want, but it must not happen because a field was left
            // blank by accident - so it is refused here and has to go through Remove instead.
            if (identities.Length == 0)
                return await FailAsync(correlationId, "ReplaceMembers", request.TaskNumber,
                    "No members given. Replacing with an empty list would remove every member - remove them individually if that is the intent.");

            // Capture what is being replaced - the legacy script wrote the old list to a report
            // first, and this is destructive.
            var previous = await GetGroupMemberNamesAsync(request.GroupIdentity, ct);
            await LogAsync(correlationId, "ReplaceMembers", "INPUT",
                $"Group={request.GroupIdentity}; Replacing {previous.Count} member(s) with {identities.Length}: {string.Join(", ", identities)}",
                request.TaskNumber, Severity.Info);
            await LogAsync(correlationId, "ReplaceMembers", "SNAP",
                $"Previous members: {(previous.Count == 0 ? "(none)" : string.Join(", ", previous))}", request.TaskNumber, Severity.Info);

            Progress(onProgress, $"Replacing membership of {request.GroupIdentity}...");
            await _host.InvokeAsync(ps => ps
                .AddCommand("Update-DistributionGroupMember")
                .AddParameter("Identity", request.GroupIdentity)
                .AddParameter("Members", identities)
                .AddParameter("BypassSecurityGroupManagerCheck", true)
                .AddParameter("Confirm", false)
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            var displayName = (await GetDetailsAsync(request.GroupIdentity, ct))?.DisplayName ?? request.GroupIdentity;
            var mailWarning = await NotifyRequesterAsync(
                correlationId, "ReplaceMembers", request.RequesterIdentity, request.TaskNumber,
                GroupNotificationMail.MembershipReplaced(displayName, identities.Length, request.TaskNumber), ct);
            if (mailWarning is not null)
                Progress(onProgress, $"WARNING: {mailWarning}");

            await LogAsync(correlationId, "ReplaceMembers", "DONE",
                $"Membership replaced with {identities.Length} member(s)", request.TaskNumber, Severity.Success);
            Progress(onProgress, "Membership replaced.");

            return new GroupOperationResult
            {
                Succeeded = true,
                DisplayName = request.GroupIdentity,
                CorrelationId = correlationId,
                Results = identities.Select(i => new MemberOperationResult(i, true, null)).ToList(),
                Warning = mailWarning,
            };
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, "ReplaceMembers", request.TaskNumber, ex.Message);
        }
    }

    private async Task<List<MemberOperationResult>> AddMembersAsync(string groupIdentity, IEnumerable<string> identities, CancellationToken ct)
    {
        var results = new List<MemberOperationResult>();

        foreach (var raw in identities)
        {
            // Other lists are referenced by name (LST.*/DST.*), people by address - completing a
            // list name with a domain would break it, so only bare user names get one.
            var resolved = LooksLikeGroupName(raw) ? raw : AppendDomain(raw);
            try
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Add-DistributionGroupMember")
                    .AddParameter("Identity", groupIdentity)
                    .AddParameter("Member", resolved)
                    .AddParameter("BypassSecurityGroupManagerCheck", true)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                results.Add(new MemberOperationResult(raw, true, null));
            }
            catch (Exception ex)
            {
                results.Add(new MemberOperationResult(raw, false, ex.Message));
            }
        }

        return results;
    }

    /// <summary>Legacy treated LST.*, DST.* and DSG.* as group references rather than mailboxes.</summary>
    private static bool LooksLikeGroupName(string identity)
        => identity.StartsWith("LST.", StringComparison.OrdinalIgnoreCase)
        || identity.StartsWith("DST.", StringComparison.OrdinalIgnoreCase)
        || identity.StartsWith("DSG.", StringComparison.OrdinalIgnoreCase);

    // ----- rename -----

    public async Task<GroupOperationResult> RenameAsync(GroupRenameRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "Rename", "STAR", "Group rename has started", request.TaskNumber, Severity.Info);

            var newName = (request.NewName ?? "").Trim();
            var newAddress = (request.NewAddress ?? "").Trim();

            if (newName.Length == 0)
                return await FailAsync(correlationId, "Rename", request.TaskNumber, "Enter the new group name.");
            if (newName.Length > NameLimit)
                return await FailAsync(correlationId, "Rename", request.TaskNumber,
                    $"The new name is {newName.Length} characters - Exchange allows at most {NameLimit}.");
            if (!newAddress.Contains('@'))
                return await FailAsync(correlationId, "Rename", request.TaskNumber, $"'{newAddress}' is not a valid e-mail address.");

            var existing = await GetDetailsAsync(request.GroupIdentity, ct);
            if (existing is null)
                return await FailAsync(correlationId, "Rename", request.TaskNumber, $"No group found for '{request.GroupIdentity}'.");

            await LogAsync(correlationId, "Rename", "INPUT",
                $"{existing.DisplayName} ({existing.PrimarySmtpAddress}) -> {newName} ({newAddress})", request.TaskNumber, Severity.Info);

            var newAlias = ExchangeNaming.ToAliasSafe(newAddress[..newAddress.IndexOf('@')]);

            // The new address is added as a secondary first, then promoted to primary - the old one
            // stays behind as an alias so mail to it keeps arriving, as the legacy script did.
            Progress(onProgress, "Renaming...");
            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-DistributionGroup")
                .AddParameter("Identity", request.GroupIdentity)
                .AddParameter("Name", newName)
                .AddParameter("DisplayName", newName)
                .AddParameter("Alias", newAlias)
                .AddParameter("EmailAddresses", new Hashtable { ["add"] = newAddress })
                .AddParameter("BypassSecurityGroupManagerCheck", true)
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-DistributionGroup")
                .AddParameter("Identity", newName)
                .AddParameter("PrimarySmtpAddress", newAddress)
                .AddParameter("BypassSecurityGroupManagerCheck", true)
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            await LogAsync(correlationId, "Rename", "RENA",
                $"Renamed to {newName} ({newAddress}); previous address {existing.PrimarySmtpAddress} kept as an alias",
                request.TaskNumber, Severity.Success);

            // Set-DistributionGroup drops Notes, so put them back (legacy re-applied them too).
            if (!string.IsNullOrWhiteSpace(existing.Notes))
            {
                try
                {
                    await _host.InvokeAsync(ps => ps
                        .AddCommand("Set-Group")
                        .AddParameter("Identity", newName)
                        .AddParameter("Notes", existing.Notes)
                        .AddParameter("ErrorAction", "Stop"), ct: ct);
                }
                catch (Exception ex)
                {
                    await LogAsync(correlationId, "Rename", "WARN", $"Could not restore the Notes: {ex.Message}", request.TaskNumber, Severity.Warning);
                }
            }

            var mailWarning = await NotifyRequesterAsync(
                correlationId, "Rename", request.RequesterIdentity, request.TaskNumber,
                GroupNotificationMail.Renamed(existing.DisplayName, newName, newAddress, existing.PrimarySmtpAddress, request.TaskNumber), ct);
            if (mailWarning is not null)
                Progress(onProgress, $"WARNING: {mailWarning}");

            await LogAsync(correlationId, "Rename", "DONE", "Group rename has completed", request.TaskNumber, Severity.Success);
            await RemoveFromCacheAsync(existing.PrimarySmtpAddress, ct);

            return new GroupOperationResult
            {
                Succeeded = true,
                DisplayName = newName,
                Address = newAddress,
                CorrelationId = correlationId,
                Warning = mailWarning,
            };
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, "Rename", request.TaskNumber, ex.Message);
        }
    }

    // ----- remove -----

    public async Task<GroupRemovalResult> RemoveAsync(GroupRemovalRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "Remove", "STAR", "Group removal has started", request.TaskNumber, Severity.Info);

            var details = await GetDetailsAsync(request.GroupIdentity, ct);
            if (details is null)
            {
                await LogAsync(correlationId, "Remove", "FAIL", $"No group found for '{request.GroupIdentity}'.", request.TaskNumber, Severity.Error);
                return GroupRemovalResult.Failed(correlationId, $"No group found for '{request.GroupIdentity}'.");
            }

            Progress(onProgress, "Capturing snapshot...");
            var snapshot = await BuildRemovalSnapshotAsync(details, ct);
            await LogAsync(correlationId, "Remove", "SNAP", snapshot, request.TaskNumber, Severity.Info);

            string removedAs;
            Progress(onProgress, $"Removing {details.DisplayName}...");

            if (details.IsUnified)
            {
                removedAs = "Unified";
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Remove-UnifiedGroup")
                    .AddParameter("Identity", request.GroupIdentity)
                    .AddParameter("Confirm", false)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
            }
            else if (details.IsDynamic)
            {
                removedAs = "Dynamic";
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Remove-DynamicDistributionGroup")
                    .AddParameter("Identity", request.GroupIdentity)
                    .AddParameter("Confirm", false)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
            }
            else
            {
                removedAs = "Distribution";
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Remove-DistributionGroup")
                    .AddParameter("Identity", request.GroupIdentity)
                    .AddParameter("BypassSecurityGroupManagerCheck", true)
                    .AddParameter("Confirm", false)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
            }

            await LogAsync(correlationId, "Remove", "REMO", $"Removed {details.DisplayName} as {removedAs}", request.TaskNumber, Severity.Success);

            var mailWarning = await NotifyRequesterAsync(
                correlationId, "Remove", request.RequesterIdentity, request.TaskNumber,
                GroupNotificationMail.Removal(details.DisplayName, request.TaskNumber), ct);
            if (mailWarning is not null)
                Progress(onProgress, $"WARNING: {mailWarning}");

            await LogAsync(correlationId, "Remove", "DONE", "Group removal has completed", request.TaskNumber, Severity.Success);
            await RemoveFromCacheAsync(details.PrimarySmtpAddress, ct);
            Progress(onProgress, $"{details.DisplayName} removed.");

            return new GroupRemovalResult
            {
                Succeeded = true,
                Snapshot = snapshot,
                RemovedAs = removedAs,
                CorrelationId = correlationId,
                Warning = mailWarning,
            };
        }
        catch (Exception ex)
        {
            await LogAsync(correlationId, "Remove", "FAIL", ex.Message, request.TaskNumber, Severity.Error);
            return GroupRemovalResult.Failed(correlationId, ex.Message);
        }
    }

    private async Task<string> BuildRemovalSnapshotAsync(GroupDetails details, CancellationToken ct)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Pre-removal snapshot for {details.DisplayName} <{details.PrimarySmtpAddress}> ({DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC)");
        sb.AppendLine($"  Type: {details.RecipientTypeDetails}");
        sb.AppendLine($"  Owners: {Join(details.Owners)}");
        sb.AppendLine($"  Authorized senders: {Join(details.AuthorizedSenders)}");
        sb.AppendLine($"  Addresses: {Join(details.Aliases)}");
        sb.AppendLine($"  RequireSenderAuthenticationEnabled: {details.RequireSenderAuthenticationEnabled}");
        sb.AppendLine($"  Notes: {details.Notes}");

        // Dynamic groups have no stored membership to capture.
        if (!details.IsDynamic)
            sb.AppendLine($"  Members: {Join(await GetGroupMemberNamesAsync(details.PrimarySmtpAddress, ct))}");

        return sb.ToString().TrimEnd();

        static string Join(IReadOnlyList<string> values) => values.Count == 0 ? "(none)" : string.Join(", ", values);
    }

    // ----- details -----

    public async Task<GroupDetails?> GetDetailsAsync(string groupIdentity, CancellationToken ct = default)
    {
        var recipient = await GetRecipientAsync(groupIdentity, ct);
        if (recipient is null)
            return null;

        var type = Str(recipient, "RecipientTypeDetails");
        var isDynamic = type.Contains("Dynamic", StringComparison.OrdinalIgnoreCase);
        var isUnified = type.Contains("GroupMailbox", StringComparison.OrdinalIgnoreCase);

        // Get-Recipient carries the type but not the settings, so fetch the real object.
        var command = isDynamic ? "Get-DynamicDistributionGroup" : isUnified ? "Get-UnifiedGroup" : "Get-DistributionGroup";
        var results = await _host.InvokeAsync(ps => ps
            .AddCommand(command)
            .AddParameter("Identity", groupIdentity)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

        if (results.Count == 0)
            return null;

        var group = results[0];

        // Unified groups report owners as ManagedBy too, so one path covers both.
        return new GroupDetails
        {
            DisplayName = Str(group, "DisplayName"),
            PrimarySmtpAddress = Str(group, "PrimarySmtpAddress"),
            RecipientTypeDetails = type,
            Owners = StrList(group, "ManagedBy"),
            AuthorizedSenders = StrList(group, "AcceptMessagesOnlyFromSendersOrMembers"),
            Aliases = StrList(group, "EmailAddresses"),
            RequireSenderAuthenticationEnabled = Bool(group, "RequireSenderAuthenticationEnabled"),
            Notes = Str(group, "Notes"),
        };
    }

    private async Task<PSObject?> GetRecipientAsync(string identity, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(identity))
            return null;

        var results = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-Recipient")
            .AddParameter("Identity", identity)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);
        return results.Count > 0 ? results[0] : null;
    }

    private async Task<List<string>> GetGroupMemberNamesAsync(string groupIdentity, CancellationToken ct)
    {
        try
        {
            var members = await _host.InvokeAsync(ps => ps
                .AddCommand("Get-DistributionGroupMember")
                .AddParameter("Identity", groupIdentity)
                .AddParameter("ResultSize", "Unlimited")
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            return members
                .Select(m =>
                {
                    var smtp = Str(m, "PrimarySmtpAddress");
                    return smtp.Length > 0 ? smtp : Str(m, "DisplayName");
                })
                .Where(s => s.Length > 0)
                .ToList();
        }
        catch
        {
            return new List<string>();
        }
    }

    public async Task<GroupOperationResult> RefreshMailTipAsync(string groupIdentity, string taskNumber, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "RefreshMailTip", "STAR", "MailTip refresh has started", taskNumber, Severity.Info);

            var details = await GetDetailsAsync(groupIdentity, ct);
            if (details is null)
                return await FailAsync(correlationId, "RefreshMailTip", taskNumber, $"Group '{groupIdentity}' was not found.");

            // The MailTip is built from ManagedBy, which is only how distribution and mail-enabled
            // security groups store ownership. Dynamic groups have no owners to read, and Microsoft
            // 365 groups keep theirs as links - rather than write a wrong tip, say so.
            if (details.IsDynamic)
                return await FailAsync(correlationId, "RefreshMailTip", taskNumber,
                    "A dynamic distribution group has no owner list to build a MailTip from.");
            if (details.IsUnified)
                return await FailAsync(correlationId, "RefreshMailTip", taskNumber,
                    "Microsoft 365 groups keep their owners as links rather than in ManagedBy - their MailTip is not maintained here.");

            var mailTip = await BuildOwnerMailTipAsync(groupIdentity, correlationId, "RefreshMailTip", taskNumber, ct);

            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-DistributionGroup")
                .AddParameter("Identity", groupIdentity)
                .AddParameter("MailTip", mailTip)
                .AddParameter("BypassSecurityGroupManagerCheck", true)
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            await LogAsync(correlationId, "RefreshMailTip", "UPGTIP", $"MailTip set to '{mailTip}'", taskNumber, Severity.Success);
            await LogAsync(correlationId, "RefreshMailTip", "DONE", "MailTip refresh has completed", taskNumber, Severity.Success);

            return new GroupOperationResult
            {
                Succeeded = true,
                DisplayName = details.DisplayName,
                Address = details.PrimarySmtpAddress,
                CorrelationId = correlationId,
                Info = mailTip,
            };
        }
        catch (Exception ex)
        {
            return await FailAsync(correlationId, "RefreshMailTip", taskNumber, ex.Message);
        }
    }

    /// <summary>"Owners: First Last, ..." from the group's current ManagedBy, capped for Exchange.</summary>
    private async Task<string> BuildOwnerMailTipAsync(
        string groupIdentity, Guid correlationId, string action, string taskNumber, CancellationToken ct)
    {
        // Expanded rather than read off the group object: ManagedBy has been observed arriving as
        // one flattened string, which then resolved to nothing and put raw values in the MailTip.
        var managedBy = await PsValues.ExpandProperty(_host, "Get-DistributionGroup", groupIdentity, "ManagedBy", ct);
        var names = new List<string>();
        var unresolved = new List<string>();

        foreach (var owner in managedBy)
        {
            // ManagedBy is a recipient name/DN, not an address - resolve it for a readable tip.
            var recipient = await GetRecipientAsync(owner, ct);
            var display = recipient is not null ? Str(recipient, "DisplayName") : "";

            if (display.Length == 0)
                unresolved.Add(owner);

            names.Add(display.Length > 0 ? display : owner.Split('/').Last());
        }

        // An owner that Exchange itself cannot look up means the value we read is not a usable
        // identity - the symptom that hid the flattened-ManagedBy bug for as long as it did.
        if (unresolved.Count > 0)
            await LogAsync(correlationId, action, "WARN",
                $"MailTip: {unresolved.Count} of {managedBy.Count} owner value(s) did not resolve to a recipient "
                + $"and were used as-is: {string.Join(" | ", unresolved)}", taskNumber, Severity.Warning);

        return Truncate("Owners: " + string.Join(", ", names), MailTipLimit);
    }

    // ----- SQL cache -----

    /// <summary>
    /// Writes the new group into dbo.Groups so the page's own search finds it right away instead of
    /// waiting for the next Import-M365GroupsToSql.ps1 run. Best-effort - the import reconciles.
    /// </summary>
    private async Task UpsertCacheAsync(string displayName, string alias, string address, GroupKind kind, string[] owners, CancellationToken ct)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return;

        try
        {
            var group = await GetRecipientAsync(address, ct);
            var externalId = group is not null ? Str(group, "ExternalDirectoryObjectId") : "";
            if (!Guid.TryParse(externalId, out var id))
                return;

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync(ct);

            const string sql = """
                MERGE dbo.Groups AS target
                USING (SELECT @Id AS Id) AS source ON target.Id = source.Id
                WHEN MATCHED THEN UPDATE SET
                    DisplayName = @DisplayName, MailNickname = @Alias, Alias = @Alias,
                    PrimarySmtpAddress = @Address, MailEnabled = 1, SecurityEnabled = @SecurityEnabled,
                    ManagedBy = @Owners, UpdatedDateTime = SYSUTCDATETIME(),
                    LastImportedAtUtc = SYSUTCDATETIME(), IsDeletedInM365 = 0, LastSeenAtUtc = SYSUTCDATETIME()
                WHEN NOT MATCHED THEN INSERT
                    (Id, DisplayName, MailNickname, Alias, PrimarySmtpAddress, GroupTypes, MailEnabled,
                     SecurityEnabled, ManagedBy, CreatedDateTime, UpdatedDateTime, LastImportedAtUtc,
                     IsDeletedInM365, LastSeenAtUtc)
                VALUES
                    (@Id, @DisplayName, @Alias, @Alias, @Address, '', 1, @SecurityEnabled, @Owners,
                     SYSUTCDATETIME(), SYSUTCDATETIME(), SYSUTCDATETIME(), 0, SYSUTCDATETIME());
                """;

            await using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = id;
            cmd.Parameters.Add("@DisplayName", SqlDbType.NVarChar, 256).Value = displayName;
            cmd.Parameters.Add("@Alias", SqlDbType.NVarChar, 128).Value = alias;
            cmd.Parameters.Add("@Address", SqlDbType.NVarChar, 256).Value = address;
            cmd.Parameters.Add("@SecurityEnabled", SqlDbType.Bit).Value = kind == GroupKind.Security;
            cmd.Parameters.Add("@Owners", SqlDbType.NVarChar, -1).Value = string.Join(";", owners);
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch
        {
            // Cache write is best-effort.
        }
    }

    private async Task RemoveFromCacheAsync(string address, CancellationToken ct)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs) || string.IsNullOrWhiteSpace(address))
            return;

        try
        {
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync(ct);
            await using var cmd = new SqlCommand("DELETE FROM dbo.Groups WHERE PrimarySmtpAddress = @Address", conn);
            cmd.Parameters.Add("@Address", SqlDbType.NVarChar, 256).Value = address;
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch
        {
            // Best-effort.
        }
    }

    // ----- helpers -----

    private static GroupOperationResult Summarize(
        Guid correlationId, string groupIdentity, IReadOnlyList<MemberOperationResult> results, string? mailWarning = null)
    {
        var failed = results.Count(r => !r.Succeeded);
        var partial = failed > 0 && failed < results.Count ? $"{failed} of {results.Count} failed." : null;
        var warning = string.Join(" ", new[] { partial, mailWarning }.Where(w => w is not null));

        return new GroupOperationResult
        {
            Succeeded = results.Any(r => r.Succeeded),
            DisplayName = groupIdentity,
            Results = results,
            CorrelationId = correlationId,
            ErrorMessage = failed == results.Count && results.Count > 0 ? "Nothing could be changed - see the results list." : null,
            Warning = warning.Length > 0 ? warning : null,
        };
    }

    // ----- confirmation e-mail -----

    /// <summary>
    /// Sends one of the <see cref="GroupNotificationMail"/> messages to the requester. Best-effort:
    /// a group that exists but whose requester wasn't told is still a success, so this never throws.
    /// Returns a warning for the UI, or null when nothing went wrong - including the "no requester
    /// was entered" case, which is a deliberate choice rather than a failure.
    /// </summary>
    private async Task<string?> NotifyRequesterAsync(
        Guid correlationId, string action, string requesterIdentity, string taskNumber,
        (string Subject, string Html) message, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(requesterIdentity))
            return null;

        try
        {
            var address = await ResolveRequesterAddressAsync(requesterIdentity, ct);
            await _mail.SendAsync(new[] { address }, message.Subject, message.Html, ct: ct);
            await LogAsync(correlationId, action, "SEND",
                $"Sent confirmation mail via {_mail.SenderDescription} to {address}", taskNumber, Severity.Success);
            return null;
        }
        catch (Exception ex)
        {
            await LogAsync(correlationId, action, "WARN",
                $"Could not send confirmation mail: {ex.Message}", taskNumber, Severity.Warning);
            return $"The confirmation e-mail to '{requesterIdentity}' could not be sent: {ex.Message}";
        }
    }

    /// <summary>
    /// Turns whatever the operator typed (SamAccountName, UPN, e-mail) into a deliverable SMTP
    /// address. Exchange is connected here and its PrimarySmtpAddress is authoritative.
    /// </summary>
    private async Task<string> ResolveRequesterAddressAsync(string identity, CancellationToken ct)
    {
        var candidate = AppendDomain(identity.Trim());

        var recipient = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-Recipient")
            .AddParameter("Identity", candidate)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

        if (recipient.Count > 0)
        {
            var smtp = Str(recipient[0], "PrimarySmtpAddress");
            if (smtp.Length > 0)
                return smtp;
        }

        if (candidate.Contains('@'))
            return candidate;

        throw new InvalidOperationException($"'{identity}' could not be resolved to an e-mail address.");
    }

    /// <summary>Owner display names for the confirmation e-mail; falls back to the identity as typed.</summary>
    private Task<List<string>> ResolveOwnerNamesAsync(IEnumerable<string> owners, CancellationToken ct)
        => RecipientLookup.ToDisplayNamesAsync(_host, owners.Select(AppendDomain), ct);

    private string AppendDomain(string identity)
    {
        var value = (identity ?? "").Trim();
        if (value.Length == 0 || value.Contains('@'))
            return value;

        var domain = _settings.Current.M365.Domain.Trim();
        return domain.Length == 0 ? value : $"{value}@{domain}";
    }

    private static IEnumerable<string> SplitIdentities(string csv)
        => (csv ?? "")
            .Replace("\r\n", ",").Replace('\n', ',')
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    private static string Truncate(string value, int max) => value.Length > max ? value[..max] : value;

    private static void Progress(Action<string>? onProgress, string message) => onProgress?.Invoke(message);

    // ----- logging -----

    private async Task<GroupOperationResult> FailAsync(Guid correlationId, string action, string taskNumber, string message)
    {
        await LogAsync(correlationId, action, "FAIL", message, taskNumber, Severity.Error);
        return GroupOperationResult.Failed(correlationId, message);
    }

    private async Task LogResultsAsync(Guid correlationId, string action, string eventCode, string label,
        IReadOnlyList<MemberOperationResult> results, string taskNumber)
    {
        foreach (var result in results)
        {
            await LogAsync(correlationId, action, result.Succeeded ? eventCode : "FAIL",
                result.Succeeded
                    ? $"Changed {label} {result.Identity}"
                    : $"Could not change {label} {result.Identity}: {result.Error}",
                taskNumber, result.Succeeded ? Severity.Success : Severity.Warning);
        }
    }

    private async Task LogAsync(Guid correlationId, string action, string eventCode, string message, string taskNumber, Severity severity)
    {
        try
        {
            await _log.WriteAsync(new LogEntry
            {
                Area = "Groups",
                Action = action,
                EventCode = eventCode,
                Severity = severity,
                Message = message,
                TaskNumber = string.IsNullOrWhiteSpace(taskNumber) ? null : taskNumber.Trim(),
                CorrelationId = correlationId,
                UserUpn = _auth.CurrentUser?.Upn,
            });
        }
        catch
        {
            // Logging is best-effort.
        }
    }

    private static string Str(PSObject? o, string name) => o?.Properties[name]?.Value?.ToString() ?? "";

    private static bool Bool(PSObject o, string name)
        => o.Properties[name]?.Value is bool b && b;

    private static List<string> StrList(PSObject o, string name) => PsValues.ToStringList(o, name);
}
