using System.Collections;
using System.Data;
using System.Management.Automation;
using System.Text;
using M365Manager.Core.Exchange;
using M365Manager.Core.M365;
using M365Manager.Core.Notifications;
using M365Manager.Core.PowerShell;
using M365Manager.Core.Settings;
using M365Manager.Data.Logging;
using M365Manager.Data.Rooms;
using Microsoft.Data.SqlClient;

namespace M365Manager.Core.RoomResources;

/// <summary>
/// See <see cref="IRoomResourceService"/>. Everything runs through <see cref="PowerShellHost"/>
/// against the same bundled Exchange Online module <see cref="ExchangeService"/> connects - rooms,
/// room lists and their security groups are all plain Exchange objects, so unlike the Teams
/// feature this needs no Graph calls and no extra Graph consent.
/// </summary>
public sealed partial class RoomResourceService : IRoomResourceService
{
    /// <summary>
    /// Booking-policy text appended to every accept/decline notification. Inlined from the legacy
    /// e:\automation\scripts\RoomResourcePolicy.htm, which lived on the old automation server.
    /// </summary>
    private const string StandardPolicyHtml = """
        The purpose of this notification is to inform you of the status of your room/resource request. If the notification indicates:<br>
        <ul type=disc>
         <li><b>Accepted</b> - you do not need to do anything further.</li>
         <li><b>Tentative</b> - please check to see if the room/resource is available at the requested time. If it is not available at the requested time, please reschedule by selecting a new time or by removing the original room/resource and selecting a new room/resource.</li>
         <li><b>Denied</b> - room is not available please find a new room. <b><font color=blue>Hint:</font></b> use the Room Finder feature to find an available room.</li>
        </ul>
        Standard Room Booking Policy is:<br>
        <ul type=disc>
         <li>Not more than {0} days into the future</li>
         <li>No meeting is longer than {1} minutes in length - <font color=blue>Hint: </font>use the Recurrence feature rather than booking a meeting that spans multiple days.</li>
        </ul>
        <p>If your request is outside of this policy the Room/Resource Delegates for this room will review and Accept or Decline your request.</p>
        """;

    /// <summary>Inlined from the legacy RoomResourceNonRecurringPolicy.htm (hoteling rooms).</summary>
    private const string HotelingPolicyHtml = """
        The purpose of this notification is to inform you of the status of your room/resource request. If the notification indicates:<br>
        <ul type=disc>
         <li><b>Accepted</b> - you do not need to do anything further.</li>
         <li><b>Denied</b> - room is not available or you are trying to book a repeating meeting, which is not allowed. <b><font color=blue>Hint:</font></b> use the Room Finder feature to find an available room.</li>
        </ul>
        Hoteling Room Booking Policy is:<br>
        <ul type=disc>
         <li>Not more than {0} days into the future</li>
         <li>No recurring, repeating or multi-day meetings allowed - <b><font color=blue>Hint:</font></b> recurrences can be removed from a meeting by editing it and selecting "Remove Recurrence" on the Recurrence tab.</li>
         <li>No meeting is longer than {1} minutes in length.</li>
        </ul>
        <p>If your request is outside of this policy then review your request and resubmit so it is within the policy.</p>
        """;

    private readonly PowerShellHost _host;
    private readonly IM365AuthService _auth;
    private readonly ISettingsService _settings;
    private readonly IRoomNamingService _naming;
    private readonly IRoomSiteRepository _sites;
    private readonly ILogService _log;
    private readonly IConnectionStringProvider _connectionStrings;
    private readonly INotificationMailService _mail;
    private readonly GraphRestClient _graph;

    public RoomResourceService(
        PowerShellHost host,
        IM365AuthService auth,
        ISettingsService settings,
        IRoomNamingService naming,
        IRoomSiteRepository sites,
        ILogService log,
        IConnectionStringProvider connectionStrings,
        INotificationMailService mail,
        GraphRestClient graph)
    {
        _host = host;
        _auth = auth;
        _settings = settings;
        _naming = naming;
        _sites = sites;
        _log = log;
        _connectionStrings = connectionStrings;
        _mail = mail;
        _graph = graph;
    }

    // ----- confirmation e-mail -----

    /// <summary>
    /// Sends one of the <see cref="RoomNotificationMail"/> messages to the requester. Best-effort:
    /// a room that exists but whose requester wasn't told is still a success, so this never throws.
    /// Returns a warning to surface in the UI, or null when nothing went wrong (including the
    /// "no requester was entered" case, which is a deliberate choice rather than a failure).
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
    /// address. Exchange is already connected here and its PrimarySmtpAddress is authoritative, so
    /// there is no need to go to Graph the way the shared mailbox path does.
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

        // Nothing found - if it at least looks like an address, let the relay have the final word.
        if (candidate.Contains('@'))
            return candidate;

        throw new InvalidOperationException($"'{identity}' could not be resolved to an e-mail address.");
    }

    // ----- preview / validation -----

    public async Task<RoomCreationPreview> PreviewAsync(RoomResourceCreationRequest request, CancellationToken ct = default)
    {
        var names = await BuildNamesAsync(request, ct);
        var errors = new List<string>();

        if (names.SiteCode.Length == 0 || names.DisplayName.Trim().Length <= names.SiteCode.Length)
            errors.Add("Enter the name as \"<SITE> <Name>\", e.g. \"NBK Meeting Room 1\".");

        if (request.Kind == ResourceKind.Room && string.IsNullOrWhiteSpace(request.Capacity))
            errors.Add("Room capacity is required.");
        else if (request.Kind == ResourceKind.Room && !int.TryParse(request.Capacity.Trim(), out _))
            errors.Add("Room capacity must be a number.");

        var site = names.SiteCode.Length > 0 ? await _sites.FindByCodeAsync(names.SiteCode, ct) : null;
        string? unknownSite = null;
        if (site is null && names.SiteCode.Length > 0)
        {
            unknownSite = names.SiteCode;
            errors.Add($"Site '{names.SiteCode}' is not maintained yet - add it under Settings > Room & resource sites (time zone + regional admin group).");
        }

        // Existence probes are best-effort: not being connected to Exchange yet shouldn't stop the
        // operator from seeing the names the convention produces. Each is probed independently -
        // one failing probe must not leave the others reading "does not exist", or CreateAsync
        // would try to create an object that is already there and abort partway, after the mailbox
        // was made.
        var mailboxExists = await TryProbeAsync("Get-Mailbox", names.Address, ct);
        var roomListExists = request.Kind == ResourceKind.Room
            ? await TryProbeAsync("Get-DistributionGroup", names.RoomListName, ct)
            : null;
        var delegateExists = await TryProbeAsync("Get-DistributionGroup", names.DelegateGroup, ct);
        var usersExists = names.UsersGroup.Length > 0
            ? await TryProbeAsync("Get-DistributionGroup", names.UsersGroup, ct)
            : null;

        if (mailboxExists == true)
            errors.Add($"A mailbox already exists at {names.Address}.");

        return new RoomCreationPreview
        {
            Names = names,
            UnknownSiteCode = unknownSite,
            TimeZone = site?.TimeZone,
            RegionalAdminGroup = site?.RegionalAdminGroup,
            MailboxExists = mailboxExists,
            RoomListExists = roomListExists,
            DelegateGroupExists = delegateExists,
            UsersGroupExists = usersExists,
            RoomListRequired = request.Kind == ResourceKind.Room,
            UsersGroupRequired = names.UsersGroup.Length > 0,
            Errors = errors,
        };
    }

    private Task<RoomNames> BuildNamesAsync(RoomResourceCreationRequest request, CancellationToken ct)
        => _naming.BuildAsync(
            request.RawName,
            request.Kind,
            request.AccessModel,
            request.Building,
            request.Floor,
            request.Capacity,
            ResolveMailboxDomain(request.MailboxDomain),
            ResolveGroupDomain(request.GroupDomain),
            ct);

    private string ResolveMailboxDomain(string overrideValue)
    {
        var value = (overrideValue ?? "").Trim();
        if (value.Length > 0)
            return value;

        var configured = _settings.Current.RoomResources.RoomMailboxDomain.Trim();
        return configured.Length > 0 ? configured : _settings.Current.M365.DefaultMailDomain.Trim();
    }

    private string ResolveGroupDomain(string overrideValue)
    {
        var value = (overrideValue ?? "").Trim();
        if (value.Length > 0)
            return value;

        var configured = _settings.Current.RoomResources.RoomGroupDomain.Trim();
        return configured.Length > 0 ? configured : _settings.Current.M365.DefaultMailDomain.Trim();
    }

    // ----- create -----

    public async Task<RoomResourceCreationResult> CreateAsync(RoomResourceCreationRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        var kindLabel = request.Kind == ResourceKind.Room ? "room" : "resource";

        try
        {
            await LogAsync(correlationId, "Create", "STAR", $"New {kindLabel} creation has started", request.TaskNumber, Severity.Info);
            Progress(onProgress, $"Validating {kindLabel} details...");

            var preview = await PreviewAsync(request, ct);
            if (!preview.CanCreate)
                return await FailCreateAsync(correlationId, request.TaskNumber, string.Join(" ", preview.Errors));

            // Stop before the mailbox is created rather than discovering halfway through that a
            // group we assumed was missing is actually there.
            if (preview.HasInconclusiveProbes)
                return await FailCreateAsync(correlationId, request.TaskNumber,
                    "Could not verify in Exchange Online which of the mailbox/room list/groups already exist. Check the Exchange connection and try again - nothing was created.");

            var names = preview.Names;
            var policy = ResolvePolicy(request.Preset, request.CustomPolicy);
            var settings = _settings.Current.RoomResources;

            await LogAsync(correlationId, "Create", "INPUT",
                $"Name={names.DisplayName}; Address={names.Address}; Kind={request.Kind}; Access={request.AccessModel}; " +
                $"Policy={request.Preset} ({policy.BookingWindowInDays}d/{policy.MaximumDurationInMinutes}min/recurring={policy.AllowRecurringMeetings}); " +
                $"Site={names.SiteCode}; TimeZone={preview.TimeZone}", request.TaskNumber, Severity.Info);

            // 1. The mailbox itself.
            Progress(onProgress, $"Creating {kindLabel} mailbox {names.DisplayName}...");
            await _host.InvokeAsync(ps =>
            {
                ps.AddCommand("New-Mailbox")
                  .AddParameter("Name", names.DisplayName)
                  .AddParameter("DisplayName", names.DisplayName)
                  .AddParameter("Alias", names.Alias)
                  .AddParameter("PrimarySmtpAddress", names.Address)
                  .AddParameter("Confirm", false)
                  .AddParameter("ErrorAction", "Stop");

                if (request.Kind == ResourceKind.Room)
                {
                    ps.AddParameter("Room", true);
                    ps.AddParameter("ResourceCapacity", int.Parse(request.Capacity.Trim()));
                }
                else
                {
                    ps.AddParameter("Equipment", true);
                }
            }, ct: ct);
            await LogAsync(correlationId, "Create", "NEW", $"{names.DisplayName} ({names.Address})", request.TaskNumber, Severity.Success);

            // 2. Wait for provisioning. Legacy spun on Get-Mailbox with no delay and no exit
            //    condition (RoomEquipNew.ps1:43-46); this is bounded, like TeamsService's wait.
            Progress(onProgress, "Waiting for mailbox provisioning to complete...");
            if (!await WaitForMailboxAsync(names.Address, onProgress, ct))
                return await FailCreateAsync(correlationId, request.TaskNumber, "Timed out waiting for the new mailbox to appear in Exchange Online.");

            // 3. Quotas and provenance stamp.
            //    No office location: the legacy RoomEquipNew.ps1 computed "Building x, Floor y" but
            //    never wrote it, and that value isn't what this organization keeps in Office
            //    fields (the site name goes there). Should it come back: Set-User asks for
            //    confirmation, so it needs -Confirm:$false - without it the hosted runspace
            //    failed with EXO's generic "A server side error has occurred".
            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-Mailbox")
                .AddParameter("Identity", names.Address)
                .AddParameter("IssueWarningQuota", "0.5GB")
                .AddParameter("ProhibitSendQuota", "0.75GB")
                .AddParameter("ProhibitSendReceiveQuota", "1.0GB")
                .AddParameter("CustomAttribute15", $"M365Manager {request.Kind} {DateTime.UtcNow:yyyy-MM-dd HH:mm} Per: {request.TaskNumber}")
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            // 4. Room list (rooms only) - created on demand, then the room joins it.
            if (request.Kind == ResourceKind.Room)
            {
                Progress(onProgress, $"Ensuring room list '{names.RoomListName}'...");
                if (preview.RoomListExists != true)
                {
                    await _host.InvokeAsync(ps =>
                    {
                        ps.AddCommand("New-DistributionGroup")
                          .AddParameter("Name", names.RoomListName)
                          .AddParameter("RoomList", true)
                          .AddParameter("ErrorAction", "Stop");
                        AddManagedBy(ps, settings.RoomListOwnerGroup);
                    }, ct: ct);
                    await LogAsync(correlationId, "Create", "NEW", $"Created room list {names.RoomListName}", request.TaskNumber, Severity.Success);
                }

                await _host.InvokeAsync(ps => ps
                    .AddCommand("Add-DistributionGroupMember")
                    .AddParameter("Identity", names.RoomListName)
                    .AddParameter("Member", names.Address)
                    .AddParameter("BypassSecurityGroupManagerCheck", true)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                await LogAsync(correlationId, "Create", "ADD", $"Added room to room list {names.RoomListName}", request.TaskNumber, Severity.Success);
            }

            // 5. Delegate group, and the authorized-users group for restricted rooms.
            Progress(onProgress, "Ensuring delegate group...");
            await EnsureAccessGroupAsync(correlationId, request.TaskNumber, names.DelegateGroup, names.DelegateGroupAddress,
                preview.DelegateGroupExists == true, request.DelegateMembers, "Room Delegates", ct);

            if (names.UsersGroup.Length > 0)
            {
                Progress(onProgress, "Ensuring authorized-users group...");
                await EnsureAccessGroupAsync(correlationId, request.TaskNumber, names.UsersGroup, names.UsersGroupAddress,
                    preview.UsersGroupExists == true, request.UserMembers, "Room Users", ct);
            }

            // 6. Calendar booking rules.
            Progress(onProgress, "Applying calendar booking policy...");
            await ApplyCalendarProcessingAsync(names.Address, request.AccessModel, request.Preset, policy,
                new[] { names.DelegateGroup }, new[] { names.UsersGroup }, ct);
            await LogAsync(correlationId, "Create", "CFG",
                $"Calendar processing: access={request.AccessModel}, policy={request.Preset}, delegates={names.DelegateGroup}" +
                (names.UsersGroup.Length > 0 ? $", users={names.UsersGroup}" : ""), request.TaskNumber, Severity.Success);

            // 7. Mailbox + calendar permissions.
            Progress(onProgress, "Granting mailbox permissions...");
            await GrantFullAccessAsync(correlationId, request.TaskNumber, names.Address, settings.GlobalAdminGroup, "global RRS admins", ct);
            await GrantFullAccessAsync(correlationId, request.TaskNumber, names.Address, names.DelegateGroup, "room delegates", ct);
            await GrantFullAccessAsync(correlationId, request.TaskNumber, names.Address, preview.RegionalAdminGroup ?? "", "regional RRS admins", ct);

            // Legacy bug (RoomEquipConfig.ps1:129): the Default/Reviewer grant used
            // ($address + $Folder) with $Folder never assigned, so it targeted the mailbox root
            // rather than the calendar. Target the calendar explicitly.
            try
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-MailboxFolderPermission")
                    .AddParameter("Identity", $"{names.Address}:\\Calendar")
                    .AddParameter("User", "Default")
                    .AddParameter("AccessRights", "Reviewer")
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                await LogAsync(correlationId, "Create", "CFG", "Granted all staff Reviewer access to the calendar", request.TaskNumber, Severity.Success);
            }
            catch (Exception ex)
            {
                await LogAsync(correlationId, "Create", "WARN", $"Could not set Default calendar permission: {ex.Message}", request.TaskNumber, Severity.Warning);
            }

            // 8. Time zone.
            if (!string.IsNullOrWhiteSpace(preview.TimeZone))
            {
                Progress(onProgress, "Setting time zone...");
                await ApplyTimeZoneAsync(names.Address, preview.TimeZone!, ct);
                await LogAsync(correlationId, "Create", "CFG", $"Time zone set to {preview.TimeZone}", request.TaskNumber, Severity.Success);
            }

            // 9. Restricted rooms advertise their delegates via a MailTip (RoomEquipConfig.ps1:23-42).
            if (request.AccessModel == RoomAccessModel.Restricted)
            {
                try
                {
                    var members = await GetGroupMembersAsync(names.DelegateGroup, ct);
                    var tip = BuildMailTip(members.Select(StripAddress).Where(n => n.Length > 0).ToList());
                    await _host.InvokeAsync(ps => ps
                        .AddCommand("Set-Mailbox")
                        .AddParameter("Identity", names.Address)
                        .AddParameter("MailTip", tip)
                        .AddParameter("ErrorAction", "Stop"), ct: ct);
                    await LogAsync(correlationId, "Create", "CFG", $"MailTip set to '{tip}'", request.TaskNumber, Severity.Success);
                }
                catch (Exception ex)
                {
                    await LogAsync(correlationId, "Create", "WARN", $"Could not set MailTip: {ex.Message}", request.TaskNumber, Severity.Warning);
                }
            }

            // The confirmation e-mail's wording depends on whether this created the site's room list
            // or joined an existing one - preview.RoomListExists was probed before anything was
            // written, so it still describes the state we found.
            string? mailWarning = null;
            if (!string.IsNullOrWhiteSpace(request.RequesterIdentity))
            {
                Progress(onProgress, $"Sending confirmation e-mail as {_mail.SenderDescription}...");

                var delegateNames = await TryGetGroupMembersAsync(names.DelegateGroup, ct);
                var userNames = request.AccessModel == RoomAccessModel.Restricted
                    ? await TryGetGroupMembersAsync(names.UsersGroup, ct)
                    : Array.Empty<string>();

                var newRoomList = preview.RoomListRequired && preview.RoomListExists != true;
                var message = RoomNotificationMail.Creation(
                    names, request.Kind, request.AccessModel, newRoomList, policy,
                    delegateNames, userNames, request.TaskNumber);

                mailWarning = await NotifyRequesterAsync(
                    correlationId, "Create", request.RequesterIdentity, request.TaskNumber, message, ct);
                if (mailWarning is not null)
                    Progress(onProgress, $"WARNING: {mailWarning}");
            }

            await LogAsync(correlationId, "Create", "DONE", $"New {kindLabel} creation has completed", request.TaskNumber, Severity.Success);
            Progress(onProgress, $"{names.DisplayName} created successfully.");

            await UpsertCacheAsync(names, request, policy, preview.TimeZone ?? "", ct);

            return new RoomResourceCreationResult
            {
                Succeeded = true,
                DisplayName = names.DisplayName,
                PrimarySmtpAddress = names.Address,
                CorrelationId = correlationId,
                WarningMessage = mailWarning,
            };
        }
        catch (Exception ex)
        {
            return await FailCreateAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    /// <summary>
    /// Creates a delegate/users security group if it isn't there yet and adds the requested members.
    /// The Notes line is built from THIS group's members - legacy (RoomEquipGroups.ps1:129) copied
    /// the delegate group's members onto the users group by mistake.
    /// </summary>
    private async Task EnsureAccessGroupAsync(
        Guid correlationId, string taskNumber, string groupName, string groupAddress,
        bool alreadyExists, string membersCsv, string notesLabel, CancellationToken ct)
    {
        var settings = _settings.Current.RoomResources;

        if (!alreadyExists)
        {
            await _host.InvokeAsync(ps =>
            {
                ps.AddCommand("New-DistributionGroup")
                  .AddParameter("Name", groupName)
                  .AddParameter("Alias", groupName)
                  .AddParameter("Type", "Security")
                  .AddParameter("ErrorAction", "Stop");
                if (groupAddress.Contains('@'))
                    ps.AddParameter("PrimarySmtpAddress", groupAddress);
                AddManagedBy(ps, settings.RoomListOwnerGroup);
            }, ct: ct);

            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-DistributionGroup")
                .AddParameter("Identity", groupName)
                .AddParameter("HiddenFromAddressListsEnabled", true)
                .AddParameter("RequireSenderAuthenticationEnabled", false)
                .AddParameter("BypassSecurityGroupManagerCheck", true)
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            await LogAsync(correlationId, "Create", "NEW", $"Created group {groupName}", taskNumber, Severity.Success);
        }
        else
        {
            await LogAsync(correlationId, "Create", "EXIST", $"Reusing existing group {groupName}", taskNumber, Severity.Info);
        }

        var added = await AddMembersAsync(groupName, membersCsv, ct);
        foreach (var result in added)
        {
            await LogAsync(correlationId, "Create", result.Succeeded ? "ADD" : "FAIL",
                result.Succeeded
                    ? $"Added {result.Identity} to {groupName}"
                    : $"Could not add {result.Identity} to {groupName}: {result.Error}",
                taskNumber, result.Succeeded ? Severity.Success : Severity.Warning);
        }

        await RefreshGroupNotesAsync(groupName, notesLabel, taskNumber, ct);
    }

    /// <summary>
    /// Rewrites a delegate/users group's Notes (the "Description" in the admin center) from its
    /// actual membership. Called after creation and after every membership change - leaving it to
    /// creation alone is what left groups describing delegates who had long since been removed.
    /// Cosmetic, so a failure here never fails the operation it follows.
    /// </summary>
    private async Task RefreshGroupNotesAsync(string groupName, string notesLabel, string taskNumber, CancellationToken ct)
    {
        try
        {
            var members = await GetGroupMembersAsync(groupName, ct);
            var names = members.Select(StripAddress).Where(n => n.Length > 0).ToList();

            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-Group")
                .AddParameter("Identity", groupName)
                .AddParameter("Notes", names.Count > 0
                    ? $"{notesLabel}: {string.Join(", ", names)} - Per: {taskNumber}"
                    : $"{notesLabel}: (none) - Per: {taskNumber}")
                .AddParameter("ErrorAction", "Stop"), ct: ct);
        }
        catch
        {
            // Notes are cosmetic - never fail the operation over them.
        }
    }

    /// <summary>GetGroupMembersAsync returns "Display Name &lt;smtp&gt;"; the address is noise in a description.</summary>
    private static string StripAddress(string member)
    {
        var bracket = member.IndexOf('<');
        return (bracket > 0 ? member[..bracket] : member).Trim();
    }

    /// <summary>Rewrites a room's MailTip from its delegate group. Best-effort, like at creation.</summary>
    private async Task RefreshRoomMailTipAsync(Guid correlationId, string roomAddress, string delegateGroup, string taskNumber, CancellationToken ct)
    {
        try
        {
            var members = await GetGroupMembersAsync(delegateGroup, ct);
            var tip = BuildMailTip(members.Select(StripAddress).Where(n => n.Length > 0).ToList());

            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-Mailbox")
                .AddParameter("Identity", roomAddress)
                .AddParameter("MailTip", tip)
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            await LogAsync(correlationId, "ChangeMembership", "UPMTIP", $"MailTip set to '{tip}'", taskNumber, Severity.Success);
        }
        catch (Exception ex)
        {
            await LogAsync(correlationId, "ChangeMembership", "WARN",
                $"Could not refresh the room MailTip: {ex.Message}", taskNumber, Severity.Warning);
        }
    }

    /// <summary>
    /// Builds and runs Set-CalendarProcessing for the access model / booking policy combination.
    /// Replaces the four near-duplicate legacy config scripts.
    /// </summary>
    private async Task ApplyCalendarProcessingAsync(
        string address, RoomAccessModel accessModel, BookingPolicyPreset preset, BookingPolicy policy,
        IReadOnlyList<string> delegateGroups, IReadOnlyList<string> usersGroups, CancellationToken ct)
    {
        var template = preset == BookingPolicyPreset.Hoteling ? HotelingPolicyHtml : StandardPolicyHtml;
        var responseText = string.Format(template, policy.BookingWindowInDays, policy.MaximumDurationInMinutes);

        var delegates = delegateGroups.Where(d => !string.IsNullOrWhiteSpace(d)).ToArray();
        var users = usersGroups.Where(u => !string.IsNullOrWhiteSpace(u)).ToArray();

        await _host.InvokeAsync(ps =>
        {
            ps.AddCommand("Set-CalendarProcessing")
              .AddParameter("Identity", address)
              .AddParameter("AutomateProcessing", "AutoAccept")
              .AddParameter("ProcessExternalMeetingMessages", false)
              .AddParameter("BookingWindowInDays", policy.BookingWindowInDays)
              .AddParameter("MaximumDurationInMinutes", policy.MaximumDurationInMinutes)
              .AddParameter("AllowRecurringMeetings", policy.AllowRecurringMeetings)
              .AddParameter("AllowConflicts", false)
              .AddParameter("MaximumConflictInstances", policy.MaximumConflictInstances)
              .AddParameter("ConflictPercentageAllowed", policy.ConflictPercentageAllowed)
              .AddParameter("AddAdditionalResponse", true)
              .AddParameter("AdditionalResponse", responseText)
              .AddParameter("DeleteComments", false)
              .AddParameter("DeleteSubject", false)
              .AddParameter("ErrorAction", "Stop");

            if (delegates.Length > 0)
                ps.AddParameter("ResourceDelegates", delegates);

            if (accessModel == RoomAccessModel.Restricted && users.Length > 0)
            {
                // Only the users group(s) may book at all; everything else is declined.
                ps.AddParameter("AllBookInPolicy", false);
                ps.AddParameter("AllRequestInPolicy", false);
                ps.AddParameter("AllRequestOutOfPolicy", false);
                ps.AddParameter("BookInPolicy", users);
                ps.AddParameter("RequestInPolicy", users);
                ps.AddParameter("RequestOutOfPolicy", users);
            }
            else
            {
                // General use: anyone books in-policy, delegates approve out-of-policy requests.
                ps.AddParameter("AllBookInPolicy", true);
                ps.AddParameter("AllRequestOutOfPolicy", true);
            }
        }, ct: ct);
    }

    private async Task ApplyTimeZoneAsync(string address, string timeZone, CancellationToken ct)
    {
        await _host.InvokeAsync(ps => ps
            .AddCommand("Set-MailboxCalendarConfiguration")
            .AddParameter("Identity", address)
            .AddParameter("WorkingHoursTimeZone", timeZone)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);

        await _host.InvokeAsync(ps => ps
            .AddCommand("Set-MailboxRegionalConfiguration")
            .AddParameter("Identity", address)
            .AddParameter("TimeZone", timeZone)
            .AddParameter("ErrorAction", "Stop"), ct: ct);
    }

    private async Task GrantFullAccessAsync(Guid correlationId, string taskNumber, string address, string trustee, string label, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(trustee))
            return;

        try
        {
            await _host.InvokeAsync(ps => ps
                .AddCommand("Add-MailboxPermission")
                .AddParameter("Identity", address)
                .AddParameter("User", trustee)
                .AddParameter("AccessRights", "FullAccess")
                .AddParameter("Confirm", false)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            await LogAsync(correlationId, "Create", "CFG", $"Granted FullAccess to {label} ({trustee})", taskNumber, Severity.Success);
        }
        catch (Exception ex)
        {
            await LogAsync(correlationId, "Create", "WARN", $"Could not grant FullAccess to {label} ({trustee}): {ex.Message}", taskNumber, Severity.Warning);
        }
    }

    // ----- update details -----

    public async Task<RoomResourceCreationResult> UpdateDetailsAsync(RoomDetailsUpdateRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "UpdateDetails", "STAR", "Room/resource detail change has started", request.TaskNumber, Severity.Info);

            var address = (request.Address ?? "").Trim();
            if (address.Length == 0)
                return await FailUpdateAsync(correlationId, request.TaskNumber, "No room/resource selected.");

            var mailbox = await GetOneAsync("Get-Mailbox", address, ct);
            if (mailbox is null)
                return await FailUpdateAsync(correlationId, request.TaskNumber, $"No mailbox found for '{address}'.");

            var changes = new List<string>();

            // The audit log wants the exact values; the confirmation e-mail goes to someone who
            // just wants to know what changed, so it gets its own plain-English list.
            var readableChanges = new List<string>();

            if (request.Kind == ResourceKind.Room && !string.IsNullOrWhiteSpace(request.Capacity))
            {
                if (!int.TryParse(request.Capacity.Trim(), out var capacity))
                    return await FailUpdateAsync(correlationId, request.TaskNumber, "Room capacity must be a number.");

                Progress(onProgress, "Updating capacity...");
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Set-Mailbox")
                    .AddParameter("Identity", address)
                    .AddParameter("ResourceCapacity", capacity)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);
                changes.Add($"capacity={capacity}");
                readableChanges.Add($"The capacity is now {capacity} people.");
            }

            if (!string.IsNullOrWhiteSpace(request.TimeZone))
            {
                Progress(onProgress, "Updating time zone...");
                await ApplyTimeZoneAsync(address, request.TimeZone.Trim(), ct);
                changes.Add($"timeZone={request.TimeZone.Trim()}");
                readableChanges.Add($"The calendar time zone is now {request.TimeZone.Trim()}.");
            }

            if (request.Preset is { } preset)
            {
                Progress(onProgress, "Updating booking policy...");
                var policy = ResolvePolicy(preset, request.CustomPolicy);

                // Reuse the room's existing delegates/users so changing the policy never silently
                // rewires who may book - that stays a membership operation.
                //
                // Set-CalendarProcessing rewrites the whole policy, so this MUST fail closed: if
                // the current settings can't be read we'd fall back to "general use", which would
                // set AllBookInPolicy=$true and silently open a restricted room to the entire
                // organization while reporting success.
                var processing = await _host.InvokeAsync(ps => ps
                    .AddCommand("Get-CalendarProcessing")
                    .AddParameter("Identity", address)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);

                if (processing.Count == 0)
                    return await FailUpdateAsync(correlationId, request.TaskNumber,
                        $"Could not read the current calendar settings for '{address}' - the booking policy was not changed.");

                // Expanded, not read off the object: both lists are handed straight back to
                // Set-CalendarProcessing, so a flattened value would rewrite the room's delegates
                // and booking permissions with one nonsense entry.
                var delegates = await PsValues.ExpandProperty(_host, "Get-CalendarProcessing", address, "ResourceDelegates", ct);
                var bookInPolicy = await PsValues.ExpandProperty(_host, "Get-CalendarProcessing", address, "BookInPolicy", ct);
                var accessModel = bookInPolicy.Count > 0 ? RoomAccessModel.Restricted : RoomAccessModel.GeneralUseSiteDelegates;

                await ApplyCalendarProcessingAsync(address, accessModel, preset, policy, delegates, bookInPolicy, ct);
                changes.Add($"policy={preset} ({policy.BookingWindowInDays}d/{policy.MaximumDurationInMinutes}min/recurring={policy.AllowRecurringMeetings})");
                readableChanges.Add(
                    $"The booking policy is now a {policy.BookingWindowInDays} day booking window, "
                    + $"no meeting longer than {policy.MaximumDurationInMinutes / 60} hr, "
                    + (policy.AllowRecurringMeetings ? "recurring meetings allowed." : "no recurring meetings."));
            }

            if (changes.Count == 0)
                return await FailUpdateAsync(correlationId, request.TaskNumber, "Nothing to change - fill in at least one field.");

            await LogAsync(correlationId, "UpdateDetails", "CHNG", $"{address}: {string.Join("; ", changes)}", request.TaskNumber, Severity.Success);

            var displayName = Str(mailbox, "DisplayName");
            var mailWarning = await NotifyRequesterAsync(
                correlationId, "UpdateDetails", request.RequesterIdentity, request.TaskNumber,
                RoomNotificationMail.DetailsChanged(displayName, readableChanges, request.TaskNumber), ct);
            if (mailWarning is not null)
                Progress(onProgress, $"WARNING: {mailWarning}");

            await LogAsync(correlationId, "UpdateDetails", "DONE", "Room/resource detail change has completed", request.TaskNumber, Severity.Success);

            await UpdateCacheDetailsAsync(address, request, ct);

            return new RoomResourceCreationResult
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
            return await FailUpdateAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    // ----- membership (the RoomResourceOOPChanges.ps1 equivalent) -----

    public async Task<RoomMembershipChangeResult> ChangeMembershipAsync(RoomMembershipChangeRequest request, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "ChangeMembership", "STAR", "Room/resource membership change has started", request.TaskNumber, Severity.Info);

            var group = (request.GroupIdentity ?? "").Trim();
            if (group.Length == 0)
            {
                return await FailMembershipAsync(correlationId, request.TaskNumber,
                    request.GroupKind == RoomGroupKind.Users
                        ? "This is a general-use room - it has no authorized-users group."
                        : "This room has no delegate group recorded.");
            }

            if (!await RecipientExistsAsync("Get-DistributionGroup", group, ct))
                return await FailMembershipAsync(correlationId, request.TaskNumber, $"Group '{group}' was not found in Exchange.");

            await LogAsync(correlationId, "ChangeMembership", "INPUT",
                $"Room={request.Address}; Group={group} ({request.GroupKind}); Mode={request.Mode}; Identities={request.Identities}",
                request.TaskNumber, Severity.Info);

            var results = request.Mode == MembershipChangeMode.Add
                ? await AddMembersAsync(group, request.Identities, ct)
                : await RemoveMembersAsync(group, request.Identities, ct);

            foreach (var result in results)
            {
                var verb = request.Mode == MembershipChangeMode.Add ? "Added" : "Removed";
                await LogAsync(correlationId, "ChangeMembership", result.Succeeded ? (request.Mode == MembershipChangeMode.Add ? "ADD" : "REMO") : "FAIL",
                    result.Succeeded
                        ? $"{verb} {result.Identity} {(request.Mode == MembershipChangeMode.Add ? "to" : "from")} {group}"
                        : $"Could not {verb.ToLowerInvariant()} {result.Identity}: {result.Error}",
                    request.TaskNumber, result.Succeeded ? Severity.Success : Severity.Warning);
            }

            var changed = results.Any(r => r.Succeeded);

            // Keep the group's description and, for a per-room delegate group, the room's MailTip in
            // step with the new membership. Creation did this; changing membership did not, so both
            // kept describing whoever was a delegate on the day the room was created.
            if (changed)
            {
                var notesLabel = request.GroupKind == RoomGroupKind.Users ? "Room Users" : "Room Delegates";
                await RefreshGroupNotesAsync(group, notesLabel, request.TaskNumber, ct);
                await LogAsync(correlationId, "ChangeMembership", "UPNOTE",
                    $"Refreshed the description of {group}", request.TaskNumber, Severity.Success);

                // A site-wide delegate group is shared by every general-use room at that site, so no
                // single room's MailTip describes it - only per-room groups get one.
                if (request.GroupKind == RoomGroupKind.Delegates && !IsSiteWideDelegateGroup(group))
                    await RefreshRoomMailTipAsync(correlationId, request.Address, group, request.TaskNumber, ct);
            }

            // Only adding people to a restricted room's authorized-users group has a template
            // (RROOPChanges.oft, "Additional Users Authorized to Reserve ..."). Removals and
            // delegate-group changes had none, and nothing was invented for them.
            var added = results.Where(r => r.Succeeded).Select(r => r.Identity).ToArray();
            string? mailWarning = null;
            if (request.GroupKind == RoomGroupKind.Users
                && request.Mode == MembershipChangeMode.Add
                && added.Length > 0)
            {
                var mailbox = await GetOneAsync("Get-Mailbox", request.Address, ct);
                var displayName = mailbox is null ? request.Address : Str(mailbox, "DisplayName");

                // The results carry the identities as typed - resolve them so the mail names people.
                var addedNames = await RecipientLookup.ToDisplayNamesAsync(_host, added, ct);

                mailWarning = await NotifyRequesterAsync(
                    correlationId, "ChangeMembership", request.RequesterIdentity, request.TaskNumber,
                    RoomNotificationMail.UsersAuthorized(displayName, addedNames, request.TaskNumber), ct);
            }

            await LogAsync(correlationId, "ChangeMembership", "DONE", "Room/resource membership change has completed", request.TaskNumber, Severity.Success);

            return new RoomMembershipChangeResult
            {
                Succeeded = results.Any(r => r.Succeeded),
                Results = results,
                CorrelationId = correlationId,
                ErrorMessage = results.All(r => !r.Succeeded) && results.Count > 0 ? "No member could be changed - see the results list." : null,
                WarningMessage = mailWarning,
            };
        }
        catch (Exception ex)
        {
            return await FailMembershipAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    public async Task<IReadOnlyList<string>> GetGroupMembersAsync(string groupIdentity, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(groupIdentity))
            return Array.Empty<string>();

        var members = await _host.InvokeAsync(ps => ps
            .AddCommand("Get-DistributionGroupMember")
            .AddParameter("Identity", groupIdentity)
            .AddParameter("ResultSize", "Unlimited")
            .AddParameter("ErrorAction", "Stop"), ct: ct);

        return members
            .Select(m =>
            {
                var display = Str(m, "DisplayName");
                var smtp = Str(m, "PrimarySmtpAddress");
                return display.Length > 0 && smtp.Length > 0 ? $"{display} <{smtp}>" : display.Length > 0 ? display : smtp;
            })
            .Where(s => s.Length > 0)
            .OrderBy(s => s)
            .ToList();
    }

    /// <summary>
    /// Group members for the confirmation e-mail. A failure here must not fail the operation the
    /// mail is about, so an unreadable group just yields no names.
    /// </summary>
    private async Task<IReadOnlyList<string>> TryGetGroupMembersAsync(string groupIdentity, CancellationToken ct)
    {
        try { return await GetGroupMembersAsync(groupIdentity, ct); }
        catch { return Array.Empty<string>(); }
    }

    private async Task<IReadOnlyList<MemberOperationResult>> AddMembersAsync(string group, string identitiesCsv, CancellationToken ct)
    {
        var results = new List<MemberOperationResult>();
        foreach (var raw in SplitIdentities(identitiesCsv))
        {
            var resolved = AppendDomain(raw);
            try
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Add-DistributionGroupMember")
                    .AddParameter("Identity", group)
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

    private async Task<IReadOnlyList<MemberOperationResult>> RemoveMembersAsync(string group, string identitiesCsv, CancellationToken ct)
    {
        var results = new List<MemberOperationResult>();
        foreach (var raw in SplitIdentities(identitiesCsv))
        {
            var resolved = AppendDomain(raw);
            try
            {
                await _host.InvokeAsync(ps => ps
                    .AddCommand("Remove-DistributionGroupMember")
                    .AddParameter("Identity", group)
                    .AddParameter("Member", resolved)
                    .AddParameter("BypassSecurityGroupManagerCheck", true)
                    .AddParameter("Confirm", false)
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

    // ----- remove -----

    public async Task<RoomRemovalResult> RemoveAsync(RoomRemovalRequest request, Action<string>? onProgress = null, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();

        try
        {
            await LogAsync(correlationId, "Remove", "STAR", "Room/resource removal has started", request.TaskNumber, Severity.Info);

            var address = (request.Address ?? "").Trim();
            if (address.Length == 0)
                return await FailRemoveAsync(correlationId, request.TaskNumber, "No room/resource selected.");

            var mailbox = await GetOneAsync("Get-Mailbox", address, ct);
            if (mailbox is null)
                return await FailRemoveAsync(correlationId, request.TaskNumber, $"No mailbox found for '{address}'.");

            // Snapshot everything BEFORE deleting - same data the legacy script wrote to a text
            // file, but into the audit log so it stays with the rest of the trail.
            Progress(onProgress, "Capturing pre-removal snapshot...");
            var snapshot = await CaptureSnapshotAsync(address, ct);
            await LogAsync(correlationId, "Remove", "SNAP", snapshot, request.TaskNumber, Severity.Info);

            var cached = await FindCachedRowAsync(address, ct);
            var removedGroups = new List<string>();

            // Drop room-list membership and per-room groups first, while the mailbox still resolves.
            if (request.RemoveOrphanedGroups && cached is not null)
            {
                if (!string.IsNullOrWhiteSpace(cached.RoomListName))
                {
                    Progress(onProgress, "Removing from room list...");
                    try
                    {
                        await _host.InvokeAsync(ps => ps
                            .AddCommand("Remove-DistributionGroupMember")
                            .AddParameter("Identity", cached.RoomListName)
                            .AddParameter("Member", address)
                            .AddParameter("BypassSecurityGroupManagerCheck", true)
                            .AddParameter("Confirm", false)
                            .AddParameter("ErrorAction", "Stop"), ct: ct);
                        await LogAsync(correlationId, "Remove", "REMO", $"Removed from room list {cached.RoomListName}", request.TaskNumber, Severity.Success);
                    }
                    catch (Exception ex)
                    {
                        await LogAsync(correlationId, "Remove", "WARN", $"Could not remove from room list: {ex.Message}", request.TaskNumber, Severity.Warning);
                    }
                }

                // Only per-room groups are deleted. A site-wide delegate group
                // (MBX.<SITE>.RRS.OutOfPolicy.DE) is shared by every general-use room at that site
                // and must survive.
                foreach (var group in new[] { cached.DelegateGroup, cached.UsersGroup })
                {
                    if (string.IsNullOrWhiteSpace(group) || IsSiteWideDelegateGroup(group))
                        continue;

                    Progress(onProgress, $"Removing group {group}...");
                    try
                    {
                        await _host.InvokeAsync(ps => ps
                            .AddCommand("Remove-DistributionGroup")
                            .AddParameter("Identity", group)
                            .AddParameter("BypassSecurityGroupManagerCheck", true)
                            .AddParameter("Confirm", false)
                            .AddParameter("ErrorAction", "Stop"), ct: ct);
                        removedGroups.Add(group);
                        await LogAsync(correlationId, "Remove", "REMO", $"Removed group {group}", request.TaskNumber, Severity.Success);
                    }
                    catch (Exception ex)
                    {
                        await LogAsync(correlationId, "Remove", "WARN", $"Could not remove group {group}: {ex.Message}", request.TaskNumber, Severity.Warning);
                    }
                }
            }

            Progress(onProgress, "Removing the mailbox...");
            await _host.InvokeAsync(ps => ps
                .AddCommand("Remove-Mailbox")
                .AddParameter("Identity", address)
                .AddParameter("Confirm", false)
                .AddParameter("ErrorAction", "Stop"), ct: ct);
            await LogAsync(correlationId, "Remove", "REMO", $"Removed mailbox {address}", request.TaskNumber, Severity.Success);

            await MarkCacheDeletedAsync(address, ct);

            // The display name was read before the mailbox was deleted - it can't be looked up now.
            var mailWarning = await NotifyRequesterAsync(
                correlationId, "Remove", request.RequesterIdentity, request.TaskNumber,
                RoomNotificationMail.Removal(Str(mailbox, "DisplayName"), request.TaskNumber), ct);
            if (mailWarning is not null)
                Progress(onProgress, $"WARNING: {mailWarning}");

            await LogAsync(correlationId, "Remove", "DONE", "Room/resource removal has completed", request.TaskNumber, Severity.Success);
            Progress(onProgress, $"{address} removed.");

            return new RoomRemovalResult
            {
                Succeeded = true,
                Snapshot = snapshot,
                RemovedGroups = removedGroups,
                CorrelationId = correlationId,
                WarningMessage = mailWarning,
            };
        }
        catch (Exception ex)
        {
            return await FailRemoveAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    /// <summary>A site-wide delegate group ends in ".RRS.OutOfPolicy.DE" - per-room ones carry the room name before ".DE".</summary>
    private static bool IsSiteWideDelegateGroup(string group)
        => group.EndsWith(".RRS.OutOfPolicy.DE", StringComparison.OrdinalIgnoreCase);

    /// <summary>Reproduces the legacy pre-removal report (RoomResourceRemove.ps1::Remove-Item).</summary>
    private async Task<string> CaptureSnapshotAsync(string address, CancellationToken ct)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Pre-removal snapshot for {address} ({DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC)");

        var sections = new (string Title, string Command)[]
        {
            ("Mailbox", "Get-Mailbox"),
            ("Mailbox statistics", "Get-MailboxStatistics"),
            ("Mailbox folder statistics", "Get-MailboxFolderStatistics"),
            ("Mailbox permissions", "Get-MailboxPermission"),
            ("Calendar processing", "Get-CalendarProcessing"),
            ("Calendar configuration", "Get-MailboxCalendarConfiguration"),
            ("Regional configuration", "Get-MailboxRegionalConfiguration"),
        };

        foreach (var (title, command) in sections)
        {
            sb.AppendLine().AppendLine($"--- {title} ---");
            try
            {
                var results = await _host.InvokeAsync(ps => ps
                    .AddCommand(command)
                    .AddParameter("Identity", address)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);

                foreach (var item in results)
                    sb.AppendLine(FormatPsObject(item, command));
            }
            catch (Exception ex)
            {
                sb.AppendLine($"(unavailable: {ex.Message})");
            }
        }

        return sb.ToString();
    }

    /// <summary>Flattens a PSObject into "Name: value" lines, trimmed to the fields worth keeping per cmdlet.</summary>
    private static string FormatPsObject(PSObject item, string command)
    {
        var interesting = command switch
        {
            "Get-Mailbox" => new[] { "DisplayName", "Name", "Alias", "PrimarySmtpAddress", "EmailAddresses", "RecipientTypeDetails", "ResourceCapacity", "Office", "WhenMailboxCreated", "ExchangeGuid", "HiddenFromAddressListsEnabled", "MailTip" },
            "Get-MailboxStatistics" => new[] { "DisplayName", "ItemCount", "TotalItemSize", "LastLogonTime" },
            "Get-MailboxFolderStatistics" => new[] { "Name", "ItemsInFolder", "FolderSize" },
            "Get-MailboxPermission" => new[] { "User", "AccessRights", "IsInherited", "Deny" },
            "Get-CalendarProcessing" => new[] { "AutomateProcessing", "AllowConflicts", "AllowRecurringMeetings", "BookingWindowInDays", "MaximumDurationInMinutes", "AllBookInPolicy", "AllRequestInPolicy", "AllRequestOutOfPolicy", "BookInPolicy", "RequestInPolicy", "RequestOutOfPolicy", "ResourceDelegates" },
            "Get-MailboxCalendarConfiguration" => new[] { "WorkingHoursTimeZone", "WorkDays", "WorkingHoursStartTime", "WorkingHoursEndTime" },
            "Get-MailboxRegionalConfiguration" => new[] { "TimeZone", "Language", "DateFormat", "TimeFormat" },
            _ => Array.Empty<string>(),
        };

        var sb = new StringBuilder();
        foreach (var name in interesting)
        {
            var property = item.Properties[name];
            if (property?.Value is null)
                continue;

            var value = property.Value is IEnumerable seq and not string
                ? string.Join(", ", seq.Cast<object?>().Where(x => x is not null))
                : property.Value.ToString();

            if (!string.IsNullOrWhiteSpace(value))
                sb.AppendLine($"  {name}: {value}");
        }
        return sb.ToString().TrimEnd();
    }

    // ----- shared Exchange helpers -----

    private async Task<bool> WaitForMailboxAsync(string address, Action<string>? onProgress, CancellationToken ct)
    {
        for (var attempt = 1; attempt <= 20; attempt++)
        {
            if (await RecipientExistsAsync("Get-Mailbox", address, ct))
                return true;

            Progress(onProgress, $"Still waiting for mailbox provisioning... ({attempt}/20)");
            await Task.Delay(TimeSpan.FromSeconds(5), ct);
        }
        return false;
    }

    private async Task<bool> RecipientExistsAsync(string command, string identity, CancellationToken ct)
        => await GetOneAsync(command, identity, ct) is not null;

    /// <summary>
    /// Existence probe that distinguishes "does not exist" (false) from "could not tell" (null),
    /// e.g. because Exchange isn't connected yet. Callers that go on to create the object must
    /// treat null as a stop signal rather than as "does not exist".
    /// </summary>
    private async Task<bool?> TryProbeAsync(string command, string identity, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(identity))
            return null;

        try
        {
            return await RecipientExistsAsync(command, identity, ct);
        }
        catch
        {
            return null;
        }
    }

    private async Task<PSObject?> GetOneAsync(string command, string identity, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(identity))
            return null;

        var results = await _host.InvokeAsync(ps => ps
            .AddCommand(command)
            .AddParameter("Identity", identity)
            .AddParameter("ErrorAction", "SilentlyContinue"), ct: ct);
        return results.Count > 0 ? results[0] : null;
    }

    private static void AddManagedBy(System.Management.Automation.PowerShell ps, string ownerGroup)
    {
        if (!string.IsNullOrWhiteSpace(ownerGroup))
            ps.AddParameter("ManagedBy", ownerGroup);
    }

    private BookingPolicy ResolvePolicy(BookingPolicyPreset preset, BookingPolicy? custom)
    {
        if (preset == BookingPolicyPreset.Custom && custom is not null)
            return custom;

        var basePolicy = BookingPolicy.For(preset);
        var settings = _settings.Current.RoomResources;

        // Settings supply the house defaults for the window/duration; the preset still decides
        // recurrence and conflict handling.
        return basePolicy with
        {
            BookingWindowInDays = settings.DefaultBookingWindowDays > 0 ? settings.DefaultBookingWindowDays : basePolicy.BookingWindowInDays,
            MaximumDurationInMinutes = settings.DefaultMaximumDurationMinutes > 0 ? settings.DefaultMaximumDurationMinutes : basePolicy.MaximumDurationInMinutes,
        };
    }

    /// <summary>A bare SamAccountName (no "@") is completed with the configured default UPN domain.</summary>
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

    private static string BuildMailTip(IReadOnlyList<string> delegateMembers)
    {
        var text = "Room Delegates: " + string.Join(", ", delegateMembers);
        return text.Length > 175 ? text[..175] : text;
    }

    // ----- SQL cache -----

    public async Task<IReadOnlyList<RoomResourceOverviewRow>> GetOverviewAsync(CancellationToken ct = default)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return Array.Empty<RoomResourceOverviewRow>();

        const string query = """
            SELECT
                ExchangeGuid, DisplayName, Alias, PrimarySmtpAddress, ResourceKind, Capacity,
                Building, Floor, Office, SiteCode, TimeZone, RoomListName, DelegateGroup, UsersGroup,
                AccessModel, BookingPolicy, BookingWindowInDays, MaximumDurationInMinutes,
                AllowRecurringMeetings, IsDeletedInM365, CreatedDateTime, LastImportedAtUtc
            FROM dbo.RoomResources
            ORDER BY DisplayName;
            """;

        await using var conn = new SqlConnection(cs);
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(query, conn);
        await using var reader = await cmd.ExecuteReaderAsync(ct);

        var list = new List<RoomResourceOverviewRow>();
        while (await reader.ReadAsync(ct))
        {
            list.Add(new RoomResourceOverviewRow
            {
                ExchangeGuid = reader.GetGuid(0).ToString(),
                DisplayName = Nvl(reader, 1),
                Alias = Nvl(reader, 2),
                PrimarySmtpAddress = Nvl(reader, 3),
                ResourceKind = Nvl(reader, 4),
                Capacity = reader.IsDBNull(5) ? null : reader.GetInt32(5),
                Building = Nvl(reader, 6),
                Floor = Nvl(reader, 7),
                Office = Nvl(reader, 8),
                SiteCode = Nvl(reader, 9),
                TimeZone = Nvl(reader, 10),
                RoomListName = Nvl(reader, 11),
                DelegateGroup = Nvl(reader, 12),
                UsersGroup = Nvl(reader, 13),
                AccessModel = Nvl(reader, 14),
                BookingPolicy = Nvl(reader, 15),
                BookingWindowInDays = reader.IsDBNull(16) ? null : reader.GetInt32(16),
                MaximumDurationInMinutes = reader.IsDBNull(17) ? null : reader.GetInt32(17),
                AllowRecurringMeetings = reader.IsDBNull(18) ? null : reader.GetBoolean(18),
                IsDeletedInM365 = !reader.IsDBNull(19) && reader.GetBoolean(19),
                CreatedDateTime = reader.IsDBNull(20) ? null : reader.GetDateTime(20),
                LastImportedAtUtc = reader.IsDBNull(21) ? default : reader.GetDateTime(21),
            });
        }
        return list;
    }

    private static string Nvl(SqlDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? "" : reader.GetString(ordinal);

    private async Task<RoomResourceOverviewRow?> FindCachedRowAsync(string address, CancellationToken ct)
    {
        var rows = await GetOverviewAsync(ct);
        return rows.FirstOrDefault(r => string.Equals(r.PrimarySmtpAddress, address, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Writes the just-created room straight into dbo.RoomResources so the grid shows it
    /// immediately. Best-effort: the scheduled import reconciles the row anyway.
    /// </summary>
    private async Task UpsertCacheAsync(RoomNames names, RoomResourceCreationRequest request, BookingPolicy policy, string timeZone, CancellationToken ct)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return;

        try
        {
            var mailbox = await GetOneAsync("Get-Mailbox", names.Address, ct);
            if (mailbox is null || !Guid.TryParse(Str(mailbox, "ExchangeGuid"), out var exchangeGuid))
                return;

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync(ct);

            const string sql = """
                MERGE dbo.RoomResources AS target
                USING (SELECT @Id AS ExchangeGuid) AS source ON target.ExchangeGuid = source.ExchangeGuid
                WHEN MATCHED THEN UPDATE SET
                    DisplayName = @DisplayName, Alias = @Alias, PrimarySmtpAddress = @Address,
                    ResourceKind = @Kind, Capacity = @Capacity, Building = @Building, Floor = @Floor,
                    Office = @Office, SiteCode = @SiteCode, TimeZone = @TimeZone,
                    RoomListName = @RoomList, DelegateGroup = @DelegateGroup, UsersGroup = @UsersGroup,
                    AccessModel = @AccessModel, BookingPolicy = @BookingPolicy,
                    BookingWindowInDays = @Window, MaximumDurationInMinutes = @Duration,
                    AllowRecurringMeetings = @Recurring, IsDeletedInM365 = 0,
                    LastImportedAtUtc = SYSUTCDATETIME(), LastSeenAtUtc = SYSUTCDATETIME()
                WHEN NOT MATCHED THEN INSERT
                    (ExchangeGuid, DisplayName, Alias, PrimarySmtpAddress, ResourceKind, Capacity,
                     Building, Floor, Office, SiteCode, TimeZone, RoomListName, DelegateGroup, UsersGroup,
                     AccessModel, BookingPolicy, BookingWindowInDays, MaximumDurationInMinutes,
                     AllowRecurringMeetings, CreatedDateTime, IsDeletedInM365, LastImportedAtUtc, LastSeenAtUtc)
                VALUES
                    (@Id, @DisplayName, @Alias, @Address, @Kind, @Capacity, @Building, @Floor, @Office,
                     @SiteCode, @TimeZone, @RoomList, @DelegateGroup, @UsersGroup, @AccessModel,
                     @BookingPolicy, @Window, @Duration, @Recurring, SYSUTCDATETIME(), 0,
                     SYSUTCDATETIME(), SYSUTCDATETIME());
                """;

            await using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@Id", SqlDbType.UniqueIdentifier).Value = exchangeGuid;
            cmd.Parameters.Add("@DisplayName", SqlDbType.NVarChar, 256).Value = names.DisplayName;
            cmd.Parameters.Add("@Alias", SqlDbType.NVarChar, 128).Value = names.Alias;
            cmd.Parameters.Add("@Address", SqlDbType.NVarChar, 256).Value = names.Address;
            cmd.Parameters.Add("@Kind", SqlDbType.NVarChar, 16).Value = request.Kind.ToString();
            cmd.Parameters.Add("@Capacity", SqlDbType.Int).Value =
                int.TryParse(request.Capacity?.Trim(), out var cap) ? cap : DBNull.Value;
            cmd.Parameters.Add("@Building", SqlDbType.NVarChar, 128).Value = NullableParam(request.Building);
            cmd.Parameters.Add("@Floor", SqlDbType.NVarChar, 32).Value = NullableParam(request.Floor);
            // Creation doesn't set the Office field, so the cache mustn't claim a value either.
            cmd.Parameters.Add("@Office", SqlDbType.NVarChar, 256).Value = DBNull.Value;
            cmd.Parameters.Add("@SiteCode", SqlDbType.NVarChar, 16).Value = NullableParam(names.SiteCode);
            cmd.Parameters.Add("@TimeZone", SqlDbType.NVarChar, 128).Value = NullableParam(timeZone);
            cmd.Parameters.Add("@RoomList", SqlDbType.NVarChar, 256).Value =
                request.Kind == ResourceKind.Room ? NullableParam(names.RoomListName) : DBNull.Value;
            cmd.Parameters.Add("@DelegateGroup", SqlDbType.NVarChar, 256).Value = NullableParam(names.DelegateGroup);
            cmd.Parameters.Add("@UsersGroup", SqlDbType.NVarChar, 256).Value = NullableParam(names.UsersGroup);
            cmd.Parameters.Add("@AccessModel", SqlDbType.NVarChar, 32).Value = request.AccessModel.ToString();
            cmd.Parameters.Add("@BookingPolicy", SqlDbType.NVarChar, 32).Value = request.Preset.ToString();
            cmd.Parameters.Add("@Window", SqlDbType.Int).Value = policy.BookingWindowInDays;
            cmd.Parameters.Add("@Duration", SqlDbType.Int).Value = policy.MaximumDurationInMinutes;
            cmd.Parameters.Add("@Recurring", SqlDbType.Bit).Value = policy.AllowRecurringMeetings;
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch
        {
            // Cache write is best-effort; the scheduled import will reconcile this row regardless.
        }
    }

    private async Task UpdateCacheDetailsAsync(string address, RoomDetailsUpdateRequest request, CancellationToken ct)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return;

        try
        {
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync(ct);

            // COALESCE keeps whatever the caller left blank untouched.
            const string sql = """
                UPDATE dbo.RoomResources SET
                    Capacity = COALESCE(@Capacity, Capacity),
                    TimeZone = COALESCE(@TimeZone, TimeZone),
                    BookingPolicy = COALESCE(@BookingPolicy, BookingPolicy),
                    BookingWindowInDays = COALESCE(@Window, BookingWindowInDays),
                    MaximumDurationInMinutes = COALESCE(@Duration, MaximumDurationInMinutes),
                    AllowRecurringMeetings = COALESCE(@Recurring, AllowRecurringMeetings),
                    LastSeenAtUtc = SYSUTCDATETIME()
                WHERE PrimarySmtpAddress = @Address;
                """;

            var policy = request.Preset is { } p ? ResolvePolicy(p, request.CustomPolicy) : null;

            await using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add("@Address", SqlDbType.NVarChar, 256).Value = address;
            cmd.Parameters.Add("@Capacity", SqlDbType.Int).Value =
                int.TryParse(request.Capacity?.Trim(), out var cap) ? cap : DBNull.Value;
            cmd.Parameters.Add("@TimeZone", SqlDbType.NVarChar, 128).Value = NullableParam(request.TimeZone);
            cmd.Parameters.Add("@BookingPolicy", SqlDbType.NVarChar, 32).Value =
                request.Preset is { } preset ? preset.ToString() : DBNull.Value;
            cmd.Parameters.Add("@Window", SqlDbType.Int).Value = policy is not null ? policy.BookingWindowInDays : DBNull.Value;
            cmd.Parameters.Add("@Duration", SqlDbType.Int).Value = policy is not null ? policy.MaximumDurationInMinutes : DBNull.Value;
            cmd.Parameters.Add("@Recurring", SqlDbType.Bit).Value = policy is not null ? policy.AllowRecurringMeetings : DBNull.Value;
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch
        {
            // best-effort, same as above
        }
    }

    private async Task MarkCacheDeletedAsync(string address, CancellationToken ct)
    {
        var cs = _connectionStrings.GetConnectionString();
        if (string.IsNullOrWhiteSpace(cs))
            return;

        try
        {
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync(ct);
            await using var cmd = new SqlCommand(
                "DELETE FROM dbo.RoomResources WHERE PrimarySmtpAddress = @Address", conn);
            cmd.Parameters.Add("@Address", SqlDbType.NVarChar, 256).Value = address;
            await cmd.ExecuteNonQueryAsync(ct);
        }
        catch
        {
            // best-effort
        }
    }

    private static object NullableParam(string? value)
        => string.IsNullOrWhiteSpace(value) ? DBNull.Value : value.Trim();

    // ----- logging -----

    private async Task<RoomResourceCreationResult> FailCreateAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, "Create", "FAIL", message, taskNumber, Severity.Error);
        return RoomResourceCreationResult.Failed(correlationId, message);
    }

    private async Task<RoomResourceCreationResult> FailUpdateAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, "UpdateDetails", "FAIL", message, taskNumber, Severity.Error);
        return RoomResourceCreationResult.Failed(correlationId, message);
    }

    private async Task<RoomMembershipChangeResult> FailMembershipAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, "ChangeMembership", "FAIL", message, taskNumber, Severity.Error);
        return RoomMembershipChangeResult.Failed(correlationId, message);
    }

    private async Task<RoomRemovalResult> FailRemoveAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, "Remove", "FAIL", message, taskNumber, Severity.Error);
        return RoomRemovalResult.Failed(correlationId, message);
    }

    private async Task LogAsync(Guid correlationId, string action, string eventCode, string message, string taskNumber, Severity severity)
    {
        try
        {
            await _log.WriteAsync(new LogEntry
            {
                Area = "RoomResources",
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
            // Logging is best-effort; never let it break the operation.
        }
    }

    private static void Progress(Action<string>? onProgress, string message) => onProgress?.Invoke(message);

    private static string Str(PSObject? o, string name) => o?.Properties[name]?.Value?.ToString() ?? "";

    private static List<string> StrList(PSObject o, string name) => PsValues.ToStringList(o, name);
}
