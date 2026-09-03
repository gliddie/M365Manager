using System.Net;

namespace M365Manager.Core.RoomResources;

/// <summary>
/// Subject and HTML body for the room/resource confirmation e-mails, following the legacy Outlook
/// templates the service desk used to send by hand (extracted to docs/mail-templates/ by
/// scripts/Export-MailTemplates.ps1):
///
///   RoomorResourceNewSite      new room list, general use
///   RestrictedRoomsNewSite     new room list, restricted
///   RoomorResourceAdditions    added to an existing room list, general use
///   RestrictedRoomAdditions    added to an existing room list, restricted
///   RoomResourceRemoval        room deleted
///   RROOPChanges               users added to a restricted room's authorized-users group
///
/// English only, like every notification this app sends. Each original left a blank line where the
/// operator pasted a screenshot of the Room Finder; the room list name goes there instead, since
/// that is what someone actually needs in order to find the rooms.
///
/// The booking-policy numbers the originals hardcoded ("90 days", "24 hr") are taken from the
/// policy actually applied. There is no legacy template for a details change - that wording is new.
/// </summary>
internal static class RoomNotificationMail
{
    public static (string Subject, string Html) Creation(
        RoomNames names,
        ResourceKind kind,
        RoomAccessModel accessModel,
        bool newRoomList,
        BookingPolicy policy,
        IReadOnlyList<string> delegateNames,
        IReadOnlyList<string> userNames,
        string taskNumber)
    {
        var restricted = accessModel == RoomAccessModel.Restricted;

        return (newRoomList, restricted) switch
        {
            (true, false) => NewSiteGeneralUse(names, policy, delegateNames, taskNumber),
            (true, true) => NewSiteRestricted(names, policy, delegateNames, userNames, taskNumber),
            (false, false) => AdditionGeneralUse(names, kind, taskNumber),
            (false, true) => AdditionRestricted(names, policy, delegateNames, userNames, taskNumber),
        };
    }

    // ----- creation, new room list -----

    private static (string, string) NewSiteGeneralUse(RoomNames names, BookingPolicy policy, IReadOnlyList<string> delegateNames, string taskNumber)
    {
        var html = $"""
            <p>Creation of the rooms for {E(names.SiteCode)} has been completed. It may take up to 72 hrs to be
            available in the offline address list. Using the Room Finder individuals can select rooms from the
            {E(names.SiteCode)} conference rooms:</p>

            {RoomListBlock(names)}

            <p>{NameList(delegateNames, "No delegate has been configured yet")} has been configured as the room
            delegate. As the delegate they will be forwarded requests that are outside of our booking policy,
            where they will be required to approve or reject those requests. Please share the information below
            with the delegate(s) so they understand the room delegate role.</p>

            {PolicyBlock(policy)}

            <p>With this information this ticket will be closed.</p>
            """;

        return ($"Conference Rooms for {names.SiteCode} - {taskNumber}", html);
    }

    private static (string, string) NewSiteRestricted(RoomNames names, BookingPolicy policy, IReadOnlyList<string> delegateNames, IReadOnlyList<string> userNames, string taskNumber)
    {
        var html = $"""
            <p>Creation of the restricted room <b>{E(names.DisplayName)}</b> has been completed. It may take up
            to 72 hrs to be available in the offline address list. Using the Room Finder individuals can select
            rooms from the {E(names.SiteCode)} restricted rooms:</p>

            {RoomListBlock(names)}

            <p>{NameList(delegateNames, "No delegate has been configured yet")} has been configured as the room
            delegate. As the delegate you will be forwarded requests that are outside of our booking policy,
            where you will be required to approve or reject those requests. The room delegate can request
            additions/changes to the restricted room delegates by submitting a service desk request.</p>

            <p>{NameList(userNames, "Nobody has been authorized yet")} has been configured, and are the only
            individuals allowed to request bookings in this room. The room delegate must request changes to the
            authorized room users group for this restricted room by submitting a service desk request.</p>

            {PolicyBlock(policy)}

            <p>With this information this ticket will be closed.</p>
            """;

        return ($"Restricted Conference Rooms for {names.SiteCode} - {taskNumber}", html);
    }

    // ----- creation, added to an existing room list -----

    private static (string, string) AdditionGeneralUse(RoomNames names, ResourceKind kind, string taskNumber)
    {
        var kindLabel = kind == ResourceKind.Room ? "room" : "resource";

        var html = $"""
            <p>The creation of the new {kindLabel} <b>{E(names.DisplayName)}</b> has been completed. It may take
            up to 72 hrs to be available in the offline address list. Using the Room Finder individuals can
            select rooms from the {E(names.SiteCode)} conference rooms:</p>

            {RoomListBlock(names)}

            <p>The existing room delegates have been designated to manage out-of-policy requests for this new
            {kindLabel}.</p>

            <p>With this information this ticket will be closed.</p>
            """;

        return ($"{names.DisplayName} Created - {taskNumber}", html);
    }

    private static (string, string) AdditionRestricted(RoomNames names, BookingPolicy policy, IReadOnlyList<string> delegateNames, IReadOnlyList<string> userNames, string taskNumber)
    {
        var hours = policy.MaximumDurationInMinutes / 60;

        var html = $"""
            <p>The creation of the new restricted room <b>{E(names.DisplayName)}</b> has been completed. It may
            take up to 72 hrs to be available in the offline address list. Using the Room Finder individuals can
            select rooms from the {E(names.SiteCode)} restricted rooms:</p>

            {RoomListBlock(names)}

            <p>The individuals configured as room delegates are responsible for managing out-of-policy requests.
            This includes meetings that are longer than {hours} hrs, more than {policy.BookingWindowInDays} days
            in the future, or meetings that conflict with existing reservations. Room delegates for this space
            are:<br>{NameList(delegateNames, "none configured yet")}</p>

            <p>Individuals who have been authorized to request reservations in this space are:<br>
            {NameList(userNames, "none configured yet")}</p>

            <p>With this information this ticket will be closed.</p>
            """;

        return ($"New Restricted Room {names.DisplayName} - {taskNumber}", html);
    }

    // ----- the other operations -----

    public static (string Subject, string Html) Removal(string displayName, string taskNumber)
    {
        var html = $"""
            <p>As requested the room <b>{E(displayName)}</b> has been removed from the system. It may take up to
            72 hrs for the removal of these items to sync to the Offline Address List. In the meanwhile, if
            individuals try to e-mail or reserve this space they will receive a delivery failure message
            indicating that the e-mail address cannot be resolved.</p>

            <p>With the removal of these items this ticket will be closed.</p>
            """;

        return ($"Removal of {displayName} - {taskNumber}", html);
    }

    public static (string Subject, string Html) UsersAuthorized(string roomDisplayName, IReadOnlyList<string> addedNames, string taskNumber)
    {
        var html = $"""
            <p>Per your request {NameList(addedNames, "the requested individuals")} have been added to the list
            of staff authorized to request reservations in <b>{E(roomDisplayName)}</b>.</p>

            <p>With this change being complete this ticket is being closed.</p>
            """;

        return ($"Additional Users Authorized to Reserve {roomDisplayName} - {taskNumber}", html);
    }

    /// <summary>No legacy template exists for this - the wording follows the others.</summary>
    public static (string Subject, string Html) DetailsChanged(string displayName, IReadOnlyList<string> changes, string taskNumber)
    {
        var list = string.Concat(changes.Select(c => $"<li>{E(c)}</li>"));

        var html = $"""
            <p>As requested the following has been changed for <b>{E(displayName)}</b>:</p>
            <ul>{list}</ul>

            <p>Calendar changes take effect immediately; any change to the room's details may take up to 72 hrs
            to sync to the Offline Address List.</p>

            <p>With this change being complete this ticket is being closed.</p>
            """;

        return ($"{displayName} Updated - {taskNumber}", html);
    }

    // ----- shared pieces -----

    /// <summary>
    /// Goes where the originals had a screenshot of the Room Finder. Equipment has no room list, so
    /// the block is dropped rather than naming something that doesn't exist.
    /// </summary>
    private static string RoomListBlock(RoomNames names) =>
        string.IsNullOrWhiteSpace(names.RoomListName)
            ? ""
            : $"""<p style="margin-left:24px"><b>{E(names.RoomListName)}</b></p>""";

    private static string PolicyBlock(BookingPolicy policy)
    {
        var hours = policy.MaximumDurationInMinutes / 60;
        var recurring = policy.AllowRecurringMeetings
            ? """
              <p>If someone is requesting a repeating event, room/resource availability is only checked for the
              first day of the meeting. If subsequent meetings have a conflict, a message will be sent to the
              requestor as well as the room delegate to remediate the conflict. If the room delegate wishes they
              can approve the conflict, but should go onto the calendar to reject/notify whoever had a previous
              reservation that the room has been allocated to another individual.</p>

              <p>We recommend that you educate your staff that if they receive a conflict they can modify the
              reservation themselves on the day of the conflict.</p>
              """
            : "<p>Recurring meetings cannot be booked in this space, and conflicting reservations are not accepted.</p>";

        return $"""
            <p><b>Rooms and Resources</b><br>
            Policies are applied to all rooms and resources. This helps to ensure effective utilization of space.</p>

            <p>The room/resource delegate will automatically receive an e-mail for requests made outside of the
            policy, to approve or reject the reservation request.</p>

            {recurring}

            <p>Our room/resource policy is:</p>
            <ul>
              <li>{policy.BookingWindowInDays} days booking window</li>
              <li>No meeting longer than {hours} hr</li>
            </ul>
            """;
    }

    /// <summary>
    /// Group members come back from Exchange as "Display Name &lt;smtp@domain&gt;"; the address is
    /// noise in a sentence, so only the name is kept.
    /// </summary>
    private static string NameList(IReadOnlyList<string> members, string whenEmpty)
    {
        var names = members
            .Select(m =>
            {
                var bracket = m.IndexOf('<');
                return (bracket > 0 ? m[..bracket] : m).Trim();
            })
            .Where(n => n.Length > 0)
            .ToArray();

        return names.Length == 0 ? E(whenEmpty) : E(string.Join(", ", names));
    }

    private static string E(string value) => WebUtility.HtmlEncode(value);
}
