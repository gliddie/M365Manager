using System.Management.Automation;
using M365Manager.Core.PowerShell;
using M365Manager.Data.Logging;

namespace M365Manager.Core.RoomResources;

/// <summary>
/// Access change - port of RoomResourceOOPChanges.ps1, the form behind the service desk's
/// "Access Change to Restricted Room" ticket: people added to or removed from a room's delegates and
/// from the list of people allowed to book it, both in one ticket and so in one run.
///
/// Legacy bugs not carried over:
/// - Add and Remove were two checkboxes on one text field per list; with both ticked only the add
///   ran and the removals were silently dropped. Here each has its own field.
/// - Spaces were stripped from the input and it was split on commas, so "Smith, John" or any display
///   name could never be used. Here a line (or ";") separates entries, and every entry is resolved to
///   exactly one recipient before anything is written - an ambiguous name is reported, not guessed.
/// - Only rooms whose calendar points at a group could be changed; rooms that list people directly in
///   ResourceDelegates/BookInPolicy (older restricted rooms) failed with "group not found". Those get
///   their calendar list rewritten instead.
/// - Removing users from a general-use room was not guarded (adding was).
/// - Changing a site-wide delegate group changed it for every general-use room at the site without a
///   word. The plan says so before anything happens.
/// - "Already a member" counted as a failure. It is reported as nothing-to-do instead.
/// </summary>
public sealed partial class RoomResourceService
{
    private const string AccessAction = "ChangeMembership";

    public async Task<RoomAccessSnapshot> GetAccessAsync(string roomAddress, CancellationToken ct = default)
    {
        var address = (roomAddress ?? "").Trim();
        if (address.Length == 0)
            throw new InvalidOperationException("No room/resource selected.");

        var mailbox = await GetOneAsync("Get-Mailbox", address, ct)
            ?? throw new InvalidOperationException($"No mailbox found for '{address}'.");

        var type = Str(mailbox, "RecipientTypeDetails");
        if (type is not ("RoomMailbox" or "EquipmentMailbox"))
            throw new InvalidOperationException($"'{Str(mailbox, "DisplayName")}' is a {type}, not a room or equipment mailbox.");

        // Read live: the SQL cache only knows the first entry of each list and is a day old.
        var delegates = await PsValues.ExpandProperty(_host, "Get-CalendarProcessing", address, "ResourceDelegates", ct);
        var bookIn = await PsValues.ExpandProperty(_host, "Get-CalendarProcessing", address, "BookInPolicy", ct);

        var delegateList = await ReadAccessListAsync(RoomAccessListKind.Delegates, delegates, ct);

        return new RoomAccessSnapshot
        {
            DisplayName = Str(mailbox, "DisplayName"),
            Address = Str(mailbox, "PrimarySmtpAddress"),
            // Same reading as LookUpForRenameAsync and the import script.
            AccessModel = bookIn.Count > 0
                ? RoomAccessModel.Restricted
                : delegateList.IsSiteWide || delegates.Count == 0
                    ? RoomAccessModel.GeneralUseSiteDelegates
                    : RoomAccessModel.GeneralUseCustomDelegates,
            Delegates = delegateList,
            BookingUsers = bookIn.Count > 0 ? await ReadAccessListAsync(RoomAccessListKind.BookingUsers, bookIn, ct) : null,
        };
    }

    public async Task<RoomAccessChangePlan> PlanAccessChangeAsync(RoomAccessChangeRequest request, CancellationToken ct = default)
    {
        var room = await GetAccessAsync(request.Address, ct);
        var items = new List<RoomAccessChangeItem>();
        var errors = new List<string>();
        var warnings = new List<string>();

        items.AddRange(await PlanListAsync(room.Delegates, request.AddDelegates, request.RemoveDelegates, ct));

        var bookingTyped = !string.IsNullOrWhiteSpace(request.AddBookingUsers) || !string.IsNullOrWhiteSpace(request.RemoveBookingUsers);
        if (bookingTyped && room.BookingUsers is null)
        {
            errors.Add($"{room.DisplayName} is a general-use room - anyone can book it, so there is no booking list to change. "
                       + "Only its delegates can be changed here.");
        }
        else if (room.BookingUsers is not null)
        {
            items.AddRange(await PlanListAsync(room.BookingUsers, request.AddBookingUsers, request.RemoveBookingUsers, ct));
        }

        foreach (var conflict in items.Where(i => i.Outcome == RoomAccessChangeOutcome.Conflict && i.IsAdd))
            errors.Add($"{conflict.Label} is to be both added to and removed from the {ListLabel(conflict.List)}.");

        // What the lists would look like afterwards - an empty one changes what the room is.
        if (RemainingCount(room.Delegates, items) == 0 && items.Any(i => i.List == RoomAccessListKind.Delegates && i.IsWork))
            warnings.Add($"This leaves {room.DisplayName} without delegates - nobody approves out-of-policy requests any more.");

        if (room.BookingUsers is { } booking
            && RemainingCount(booking, items) == 0
            && items.Any(i => i.List == RoomAccessListKind.BookingUsers && i.IsWork))
        {
            errors.Add($"This would leave nobody authorized to book {room.DisplayName}. "
                       + "Opening the room to everyone is an access-model change, not an access change.");
        }

        if (room.Delegates.IsSiteWide && items.Any(i => i.List == RoomAccessListKind.Delegates && i.IsWork))
            warnings.Add($"{room.Delegates.Source} is the site-wide delegate group - the change applies to every general-use room at the site, not only to {room.DisplayName}.");

        if (room.BookingSharesDelegateGroup && items.Any(i => i.IsWork))
            warnings.Add($"Booking rights on {room.DisplayName} come from its delegate group - a change to one list changes the other too.");

        return new RoomAccessChangePlan
        {
            Room = room,
            Request = request,
            Items = items,
            Errors = errors,
            Warnings = warnings,
        };
    }

    public async Task<RoomAccessChangeResult> ApplyAccessChangeAsync(RoomAccessChangePlan plan, CancellationToken ct = default)
    {
        var correlationId = Guid.NewGuid();
        var request = plan.Request;
        var room = plan.Room;

        try
        {
            await LogAsync(correlationId, AccessAction, "STAR", "Room/resource access change has started", request.TaskNumber, Severity.Info);

            if (plan.Errors.Count > 0)
                return await FailAccessAsync(correlationId, request.TaskNumber, string.Join(" ", plan.Errors));
            if (!plan.HasWork)
                return await FailAccessAsync(correlationId, request.TaskNumber, "Nothing to change - see the list of skipped entries.");

            await LogAsync(correlationId, AccessAction, "INPUT",
                $"Room={room.Address}; Delegates={room.Delegates.Source}; Booking={room.BookingUsers?.Source ?? "(general use)"}; "
                + $"AddDelegates={OneLine(request.AddDelegates)}; RemoveDelegates={OneLine(request.RemoveDelegates)}; "
                + $"AddBooking={OneLine(request.AddBookingUsers)}; RemoveBooking={OneLine(request.RemoveBookingUsers)}",
                request.TaskNumber, Severity.Info);

            foreach (var skipped in plan.Skipped)
            {
                await LogAsync(correlationId, AccessAction, "SKIP",
                    $"{(skipped.IsAdd ? "Add" : "Remove")} {skipped.Label} ({ListLabel(skipped.List)}): {skipped.Outcome.Describe()}",
                    request.TaskNumber, Severity.Warning);
            }

            var results = new List<RoomAccessChangeItemResult>();
            results.AddRange(await ApplyListAsync(room, room.Delegates, plan.Items, ct));
            if (room.BookingUsers is not null)
                results.AddRange(await ApplyListAsync(room, room.BookingUsers, plan.Items, ct));

            foreach (var result in results)
            {
                var item = result.Item;
                await LogAsync(correlationId, AccessAction,
                    result.Succeeded ? (item.IsAdd ? "ADD" : "REMO") : "FAIL",
                    result.Succeeded
                        ? $"{(item.IsAdd ? "Added" : "Removed")} {item.Label} {(item.IsAdd ? "to" : "from")} the {ListLabel(item.List)} ({SourceOf(room, item.List)})"
                        : $"Could not {(item.IsAdd ? "add" : "remove")} {item.Label}: {result.Error}",
                    request.TaskNumber, result.Succeeded ? Severity.Success : Severity.Warning);
            }

            var done = results.Where(r => r.Succeeded).Select(r => r.Item).ToList();

            // Keep the group descriptions and the MailTip in step with who is actually on the lists.
            if (done.Any(i => i.List == RoomAccessListKind.Delegates))
            {
                if (room.Delegates.Storage == RoomAccessStorage.Group)
                    await RefreshGroupNotesAsync(room.Delegates.GroupIdentity, "Room Delegates", request.TaskNumber, ct);

                // Restricted rooms advertise their delegates in a MailTip (RoomEquipConfig.ps1:23-42).
                // A site-wide group belongs to no single room.
                if (room.AccessModel == RoomAccessModel.Restricted && !room.Delegates.IsSiteWide)
                    await RefreshRoomMailTipAsync(correlationId, room.Address, await CurrentNamesAsync(room, room.Delegates, ct), request.TaskNumber, ct);
            }

            if (done.Any(i => i.List == RoomAccessListKind.BookingUsers)
                && room.BookingUsers is { Storage: RoomAccessStorage.Group } users
                && !room.BookingSharesDelegateGroup)
            {
                await RefreshGroupNotesAsync(users.GroupIdentity, "Room Users", request.TaskNumber, ct);
            }

            string? mailWarning = null;
            if (done.Count > 0)
            {
                // Display names, not what was typed - "118378" means nothing to the requester.
                static IReadOnlyList<string> Names(IEnumerable<RoomAccessChangeItem> items) =>
                    items.Select(i => i.Member?.DisplayName ?? i.Typed).ToList();

                mailWarning = await NotifyRequesterAsync(correlationId, AccessAction, request.RequesterIdentity, request.TaskNumber,
                    RoomNotificationMail.AccessChanged(room.DisplayName,
                        addedUsers: Names(done.Where(i => i.List == RoomAccessListKind.BookingUsers && i.IsAdd)),
                        removedUsers: Names(done.Where(i => i.List == RoomAccessListKind.BookingUsers && !i.IsAdd)),
                        addedDelegates: Names(done.Where(i => i.List == RoomAccessListKind.Delegates && i.IsAdd)),
                        removedDelegates: Names(done.Where(i => i.List == RoomAccessListKind.Delegates && !i.IsAdd)),
                        request.TaskNumber), ct);
            }

            var failed = results.Count(r => !r.Succeeded);
            await LogAsync(correlationId, AccessAction, "DONE",
                $"Room/resource access change has completed ({done.Count} applied, {failed} failed)",
                request.TaskNumber, failed == 0 ? Severity.Success : Severity.Warning);

            return new RoomAccessChangeResult
            {
                Succeeded = done.Count > 0,
                Results = results,
                CorrelationId = correlationId,
                ErrorMessage = done.Count == 0 ? "No change could be applied - see the results." : null,
                WarningMessage = mailWarning,
            };
        }
        catch (Exception ex)
        {
            return await FailAccessAsync(correlationId, request.TaskNumber, ex.Message);
        }
    }

    // ----- reading -----

    /// <summary>
    /// One entry that is a group means the room is managed through that group's membership.
    /// Anything else - people, several entries, nothing - is a list kept on the calendar itself.
    /// </summary>
    private async Task<RoomAccessList> ReadAccessListAsync(RoomAccessListKind kind, IReadOnlyList<string> entries, CancellationToken ct)
    {
        if (entries.Count == 1)
        {
            var recipient = await GetOneAsync("Get-Recipient", entries[0], ct);
            if (recipient is not null && Str(recipient, "RecipientTypeDetails").Contains("Group", StringComparison.OrdinalIgnoreCase))
            {
                var name = Str(recipient, "Name");
                var smtp = Str(recipient, "PrimarySmtpAddress");
                var identity = smtp.Length > 0 ? smtp : entries[0];

                var members = await _host.InvokeAsync(ps => ps
                    .AddCommand("Get-DistributionGroupMember")
                    .AddParameter("Identity", identity)
                    .AddParameter("ResultSize", "Unlimited")
                    .AddParameter("ErrorAction", "Stop"), ct: ct);

                return new RoomAccessList
                {
                    Kind = kind,
                    Storage = RoomAccessStorage.Group,
                    GroupIdentity = identity,
                    GroupName = name,
                    IsSiteWide = IsSiteWideDelegateGroup(name) || IsSiteWideDelegateGroup(entries[0]),
                    Members = members.Select(ToAccessMember).OrderBy(m => m.DisplayName, StringComparer.OrdinalIgnoreCase).ToList(),
                };
            }
        }

        var listed = new List<RoomAccessMember>();
        foreach (var entry in entries)
        {
            // BookInPolicy hands back legacyExchangeDNs. An entry that no longer resolves (a deleted
            // user) is kept as-is, so it can still be removed and is not silently dropped on rewrite.
            var recipient = await GetOneAsync("Get-Recipient", entry, ct);
            listed.Add(recipient is null ? new RoomAccessMember(entry, entry, "", null) : ToAccessMember(recipient));
        }

        return new RoomAccessList
        {
            Kind = kind,
            Storage = RoomAccessStorage.CalendarEntries,
            IsSiteWide = entries.Any(IsSiteWideDelegateGroup),
            Members = listed,
        };
    }

    private static RoomAccessMember ToAccessMember(PSObject recipient)
    {
        var display = Str(recipient, "DisplayName");
        var smtp = Str(recipient, "PrimarySmtpAddress");
        var guid = Guid.TryParse(Str(recipient, "Guid"), out var g) ? g : (Guid?)null;
        var identity = smtp.Length > 0 ? smtp : guid?.ToString() ?? Str(recipient, "Identity");
        return new RoomAccessMember(identity, display.Length > 0 ? display : identity, smtp, guid);
    }

    // ----- planning -----

    private async Task<List<RoomAccessChangeItem>> PlanListAsync(RoomAccessList list, string addText, string removeText, CancellationToken ct)
    {
        var adds = new List<RoomAccessChangeItem>();
        var removes = new List<RoomAccessChangeItem>();

        foreach (var typed in SplitAccessIdentities(addText))
        {
            var (member, outcome) = await ResolveTypedAsync(typed, ct);
            if (member is not null)
                outcome = FindOnList(list, member) is not null ? RoomAccessChangeOutcome.AlreadyOnList : RoomAccessChangeOutcome.WillAdd;

            if (member is null || !adds.Any(a => SameMember(a.Member, member)))
                adds.Add(new RoomAccessChangeItem(list.Kind, true, typed, member, outcome));
        }

        foreach (var typed in SplitAccessIdentities(removeText))
        {
            var (member, outcome) = await ResolveTypedAsync(typed, ct);

            // A stale entry (deleted user, raw legacyExchangeDN) no longer resolves, but it is on the
            // list under exactly the text shown - so it can still be removed.
            var onList = member is not null ? FindOnList(list, member) : FindOnListByText(list, typed);
            if (onList is not null)
                (member, outcome) = (onList, RoomAccessChangeOutcome.WillRemove);
            else if (member is not null)
                outcome = RoomAccessChangeOutcome.NotOnList;

            if (member is null || !removes.Any(r => SameMember(r.Member, member)))
                removes.Add(new RoomAccessChangeItem(list.Kind, false, typed, member, outcome));
        }

        // The same person on both sides of one list is a typo in the ticket, not something to guess at.
        for (var i = 0; i < adds.Count; i++)
        {
            if (adds[i].Member is null || !removes.Any(r => SameMember(r.Member, adds[i].Member)))
                continue;

            adds[i] = adds[i] with { Outcome = RoomAccessChangeOutcome.Conflict };
            for (var j = 0; j < removes.Count; j++)
            {
                if (SameMember(removes[j].Member, adds[i].Member))
                    removes[j] = removes[j] with { Outcome = RoomAccessChangeOutcome.Conflict };
            }
        }

        return adds.Concat(removes).ToList();
    }

    /// <summary>
    /// Exactly one recipient, or none. Exchange answers a name that matches several recipients with
    /// an error rather than a list, so the error text is what tells "ambiguous" from "unknown".
    /// </summary>
    private async Task<(RoomAccessMember? Member, RoomAccessChangeOutcome Outcome)> ResolveTypedAsync(string typed, CancellationToken ct)
    {
        var value = AddressInBrackets(typed) ?? typed;
        var candidate = AppendDomain(value);
        var ambiguous = false;

        foreach (var attempt in value == candidate ? new[] { value } : new[] { value, candidate })
        {
            try
            {
                var found = await _host.InvokeAsync(ps => ps
                    .AddCommand("Get-Recipient")
                    .AddParameter("Identity", attempt)
                    .AddParameter("ErrorAction", "Stop"), ct: ct);

                if (found.Count == 1)
                    return (ToAccessMember(found[0]), RoomAccessChangeOutcome.WillAdd);
                if (found.Count > 1)
                    ambiguous = true;
            }
            catch (Exception ex) when (ex.Message.Contains("multiple", StringComparison.OrdinalIgnoreCase))
            {
                ambiguous = true;
            }
            catch (Exception ex) when (ex.Message.Contains("couldn't be found", StringComparison.OrdinalIgnoreCase)
                                       || ex.Message.Contains("could not be found", StringComparison.OrdinalIgnoreCase)
                                       || ex.Message.Contains("not found", StringComparison.OrdinalIgnoreCase))
            {
                // Try the next spelling.
            }
        }

        return (null, ambiguous ? RoomAccessChangeOutcome.Ambiguous : RoomAccessChangeOutcome.NotFound);
    }

    /// <summary>"Smith, John &lt;john.smith@contoso.com&gt;" - the form the current-members list shows, so a line can be copied over.</summary>
    private static string? AddressInBrackets(string typed)
    {
        var open = typed.LastIndexOf('<');
        var close = typed.LastIndexOf('>');
        return open >= 0 && close > open + 1 ? typed[(open + 1)..close].Trim() : null;
    }

    /// <summary>
    /// One entry per line or per ";". Only when neither is used do commas separate - otherwise a
    /// "Last, First" display name would be cut in two, which is exactly what the legacy form did.
    /// </summary>
    internal static IReadOnlyList<string> SplitAccessIdentities(string? text)
    {
        var value = text ?? "";
        var separators = value.IndexOfAny(new[] { '\n', ';' }) >= 0 ? new[] { '\r', '\n', ';' } : new[] { ',' };

        return value
            .Split(separators, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(s => s.TrimStart('-', '*', '•', ' ', '\t').Trim())
            .Where(s => s.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static RoomAccessMember? FindOnList(RoomAccessList list, RoomAccessMember member)
        => list.Members.FirstOrDefault(m => SameMember(m, member));

    private static RoomAccessMember? FindOnListByText(RoomAccessList list, string typed)
    {
        var value = AddressInBrackets(typed) ?? typed;
        return list.Members.FirstOrDefault(m =>
            string.Equals(m.Identity, value, StringComparison.OrdinalIgnoreCase)
            || string.Equals(m.DisplayName, typed, StringComparison.OrdinalIgnoreCase)
            || string.Equals(m.Label, typed, StringComparison.OrdinalIgnoreCase));
    }

    private static bool SameMember(RoomAccessMember? a, RoomAccessMember? b)
    {
        if (a is null || b is null)
            return false;
        if (a.ObjectGuid is { } ga && b.ObjectGuid is { } gb)
            return ga == gb;
        return (a.Address.Length > 0 && string.Equals(a.Address, b.Address, StringComparison.OrdinalIgnoreCase))
               || string.Equals(a.Identity, b.Identity, StringComparison.OrdinalIgnoreCase);
    }

    private static int RemainingCount(RoomAccessList list, IReadOnlyList<RoomAccessChangeItem> items)
    {
        var mine = items.Where(i => i.List == list.Kind).ToList();
        return list.Members.Count
               - mine.Count(i => i.Outcome == RoomAccessChangeOutcome.WillRemove)
               + mine.Count(i => i.Outcome == RoomAccessChangeOutcome.WillAdd);
    }

    // ----- applying -----

    private async Task<List<RoomAccessChangeItemResult>> ApplyListAsync(
        RoomAccessSnapshot room, RoomAccessList list, IReadOnlyList<RoomAccessChangeItem> allItems, CancellationToken ct)
    {
        var items = allItems.Where(i => i.List == list.Kind && i.IsWork && i.Member is not null).ToList();
        var results = new List<RoomAccessChangeItemResult>();
        if (items.Count == 0)
            return results;

        if (list.Storage == RoomAccessStorage.Group)
        {
            foreach (var item in items)
            {
                try
                {
                    await _host.InvokeAsync(ps =>
                    {
                        ps.AddCommand(item.IsAdd ? "Add-DistributionGroupMember" : "Remove-DistributionGroupMember")
                          .AddParameter("Identity", list.GroupIdentity)
                          .AddParameter("Member", item.Member!.Identity)
                          .AddParameter("BypassSecurityGroupManagerCheck", true)
                          .AddParameter("ErrorAction", "Stop");
                        if (!item.IsAdd)
                            ps.AddParameter("Confirm", false);
                    }, ct: ct);
                    results.Add(new RoomAccessChangeItemResult(item, true, null));
                }
                catch (Exception ex) when (IsAlreadyInState(ex))
                {
                    // Someone got there between plan and apply - the outcome is what was asked for.
                    results.Add(new RoomAccessChangeItemResult(item, true, null));
                }
                catch (Exception ex)
                {
                    results.Add(new RoomAccessChangeItemResult(item, false, ex.Message));
                }
            }

            return results;
        }

        // A calendar list is one property - rewritten in one call, so it succeeds or fails as a whole.
        var remaining = list.Members
            .Where(m => !items.Any(i => !i.IsAdd && SameMember(i.Member, m)))
            .Select(m => m.Identity)
            .Concat(items.Where(i => i.IsAdd).Select(i => i.Member!.Identity))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        try
        {
            var parameter = list.Kind == RoomAccessListKind.Delegates ? "ResourceDelegates" : "BookInPolicy";
            await _host.InvokeAsync(ps => ps
                .AddCommand("Set-CalendarProcessing")
                .AddParameter("Identity", room.Address)
                .AddParameter(parameter, remaining.Length > 0 ? remaining : null)
                .AddParameter("ErrorAction", "Stop"), ct: ct);

            results.AddRange(items.Select(i => new RoomAccessChangeItemResult(i, true, null)));
        }
        catch (Exception ex)
        {
            results.AddRange(items.Select(i => new RoomAccessChangeItemResult(i, false, ex.Message)));
        }

        return results;
    }

    private static bool IsAlreadyInState(Exception ex)
        => ex.Message.Contains("already a member", StringComparison.OrdinalIgnoreCase)
           || ex.Message.Contains("isn't a member", StringComparison.OrdinalIgnoreCase)
           || ex.Message.Contains("is not a member", StringComparison.OrdinalIgnoreCase);

    /// <summary>Names on a list after the change, read back rather than computed - for the MailTip.</summary>
    private async Task<IReadOnlyList<string>> CurrentNamesAsync(RoomAccessSnapshot room, RoomAccessList list, CancellationToken ct)
    {
        var fresh = list.Storage == RoomAccessStorage.Group
            ? await ReadAccessListAsync(list.Kind, new[] { list.GroupIdentity }, ct)
            : (await GetAccessAsync(room.Address, ct)).Delegates;
        return fresh.Members.Select(m => m.DisplayName).ToList();
    }

    private static string SourceOf(RoomAccessSnapshot room, RoomAccessListKind kind)
        => kind == RoomAccessListKind.Delegates ? room.Delegates.Source : room.BookingUsers?.Source ?? "";

    internal static string ListLabel(RoomAccessListKind kind)
        => kind == RoomAccessListKind.Delegates ? "room delegates" : "people authorized to book";

    private static string OneLine(string text)
    {
        var parts = SplitAccessIdentities(text);
        return parts.Count == 0 ? "-" : string.Join(" | ", parts);
    }

    private async Task<RoomAccessChangeResult> FailAccessAsync(Guid correlationId, string taskNumber, string message)
    {
        await LogAsync(correlationId, AccessAction, "FAIL", message, taskNumber, Severity.Error);
        return RoomAccessChangeResult.Failed(correlationId, message);
    }
}
