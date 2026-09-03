using System.Net;
using M365Manager.Core.Settings;

namespace M365Manager.Core.Groups;

/// <summary>
/// Subject and HTML body for the group confirmation e-mails, following the legacy Outlook templates
/// (extracted to docs/mail-templates/ by scripts/Export-MailTemplates.ps1):
///
///   DLNew                            a distribution/security group was created
///   DLRenamed                        a group was renamed
///   DLRemoval                        a group was deleted
///   DistributionListOwnershipChange  owners added/removed/replaced
///   DLMembershipReplace              the whole member list was replaced
///   DLRestrictedAccessChange         authorized senders added/removed
///
/// English only, like every notification this app sends. DLNew and DLRenamed left a blank line
/// after "you can select this from the Global Address List:" where the operator pasted a
/// screenshot; the group's address goes there instead - that is what someone needs to find it.
///
/// Not used: DLRestrictedAccessGranted, which is the same message as DLRestrictedAccessChange but
/// addressed to the people who were granted access ("you have been given access") rather than to
/// the requester. Everything here goes to the requester, so the "Change" variant is the fitting one.
/// Alias changes have no legacy template and send nothing.
/// </summary>
internal static class GroupNotificationMail
{
    public static (string Subject, string Html) Creation(
        string displayName,
        string address,
        IReadOnlyList<string> aliases,
        IReadOnlyList<string> ownerNames,
        bool allowExternalSenders,
        SharedMailboxSettings settings,
        string taskNumber)
    {
        // The original states an "additional address" unconditionally; it only makes sense when an
        // alias was actually added.
        var aliasSentence = aliases.Count == 0
            ? ""
            : $" An additional address of <b>{E(string.Join(", ", aliases))}</b> has been configured to receive e-mail for this group.";

        var sendersSentence = allowExternalSenders
            ? "configured to allow e-mail from any source to be sent to this group"
            : "configured to accept e-mail from internal senders only";

        var html = $"""
            <p>Per your request the distribution list named <b>{E(displayName)}</b> has been created and
            {sendersSentence}.{aliasSentence} It may take up to 72 hrs for this new list to sync to the offline
            address list; in the meanwhile you can select it from the Global Address List:</p>

            {AddressBlock(address)}

            <p>{NameList(ownerNames, "No owner has been recorded")} has/have been made the owner(s) of the
            distribution list and will be responsible for updating the group membership. Below is a link to the
            document that provides instructions on how to modify the group membership.</p>

            {MembershipLink(settings)}

            <p>If this distribution group is no longer needed please contact the Service Desk and have a request
            created to have the list removed from the system.</p>

            <p>With this information your ticket is being closed.</p>
            """;

        return ($"{displayName} List Created - {taskNumber}", html);
    }

    /// <summary>
    /// A dynamic group has no owners and no member list to maintain, so the membership paragraph and
    /// its help link are dropped - the whole point is that Exchange keeps the membership current.
    /// </summary>
    public static (string Subject, string Html) DynamicCreation(string displayName, string address, string taskNumber)
    {
        var html = $"""
            <p>Per your request the dynamic distribution list named <b>{E(displayName)}</b> has been created. Its
            membership is evaluated by Exchange every time a message is sent to it, so there is no member list to
            maintain. It may take up to 72 hrs for this new list to sync to the offline address list; in the
            meanwhile you can select it from the Global Address List:</p>

            {AddressBlock(address)}

            <p>If the conditions that determine who belongs to this list need to change, or the list is no longer
            needed, please contact the Service Desk and have a request created.</p>

            <p>With this information your ticket is being closed.</p>
            """;

        return ($"{displayName} List Created - {taskNumber}", html);
    }

    public static (string Subject, string Html) Renamed(string oldName, string newName, string newAddress, string oldAddress, string taskNumber)
    {
        var html = $"""
            <p>Per your request the distribution list named <b>{E(oldName)}</b> has been renamed to
            <b>{E(newName)}</b>. It may take up to 72 hrs for this change to sync to the offline address list; in
            the meanwhile you can select it from the Global Address List:</p>

            {AddressBlock(newAddress)}

            <p>The previous address <b>{E(oldAddress)}</b> has been kept as an alias, so mail sent to it will
            still reach the group.</p>

            <p>No changes to the ownership of the distribution list have been made.</p>

            <p>If this distribution group is no longer needed please contact the Service Desk and have a request
            created to have the list removed from the system.</p>

            <p>With this information this ticket will be closed.</p>
            """;

        return ($"{oldName} Renamed - {taskNumber}", html);
    }

    public static (string Subject, string Html) Removal(string displayName, string taskNumber)
    {
        var html = $"""
            <p>As requested the distribution list <b>{E(displayName)}</b> has been removed from the system. It may
            take up to 72 hrs for the removal of these items to sync to the Offline Address List. In the meanwhile,
            if individuals try to e-mail this list they will receive a delivery failure message indicating that the
            e-mail address cannot be resolved.</p>

            <p>With the removal of these items this ticket is being closed.</p>
            """;

        return ($"Removal of {displayName} Distribution List - {taskNumber}", html);
    }

    /// <summary>
    /// The original was a "Select one:" template holding four alternative paragraphs plus an
    /// editorial note, which the operator picked from by hand. The mode picks instead, and the
    /// membership link is dropped for Microsoft 365 groups - exactly what the note asked for
    /// ("Keep this ... if this is not a unified group"), since their membership is managed in
    /// Teams/Outlook rather than through that article.
    /// </summary>
    public static (string Subject, string Html) OwnershipChanged(
        string displayName,
        ListChangeMode mode,
        IReadOnlyList<string> changedNames,
        bool isUnifiedGroup,
        SharedMailboxSettings settings,
        string taskNumber)
    {
        var kind = isUnifiedGroup ? "group" : "distribution list";
        var names = NameList(changedNames, "the requested individuals");

        var opening = mode switch
        {
            ListChangeMode.Add =>
                $"Per the subject Service Desk ticket {names} has/have been added as an owner of the "
                + $"<b>{E(displayName)}</b> {kind}. Ownership grants the rights to update the membership of the group.",
            ListChangeMode.Remove =>
                $"Per the subject Service Desk ticket {names} has/have been removed as an owner of the "
                + $"<b>{E(displayName)}</b> {kind}. The remaining owners keep the rights to update the membership "
                + "of the group.",
            _ =>
                $"Per the subject Service Desk ticket the ownership of the <b>{E(displayName)}</b> {kind} has been "
                + $"transferred. {names} is/are now the owner(s). Ownership grants the rights to update the "
                + "membership of the group.",
        };

        var html = $"""
            <p>{opening}</p>

            {(isUnifiedGroup ? "" : MembershipLink(settings))}

            <p>With this information this ticket will be closed.</p>
            """;

        return ($"{displayName} Ownership Change - {taskNumber}", html);
    }

    public static (string Subject, string Html) MembershipReplaced(string displayName, int memberCount, string taskNumber)
    {
        var html = $"""
            <p>As requested the membership of the group <b>{E(displayName)}</b> has been updated with the new list
            of members provided within your request. It now has {memberCount} member(s).</p>

            <p>With the requested changes complete we are closing the subject ticket.</p>
            """;

        return ($"Group Membership Replaced for {displayName} - {taskNumber}", html);
    }

    public static (string Subject, string Html) AuthorizedSendersChanged(
        string displayName, ListChangeMode mode, IReadOnlyList<string> changedNames, string taskNumber)
    {
        var verb = mode == ListChangeMode.Remove ? "had access removed to" : "been given access to";
        var html = $"""
            <p>As requested and with the appropriate authorization {NameList(changedNames, "the requested individuals")}
            have {verb} use the <b>{E(displayName)}</b> e-mail group. It may take 15-30 minutes for these access
            changes to synchronize through the environment.</p>

            <p>With this change complete the subject ticket is being closed.</p>
            """;

        return ($"Access to Restricted Email Group {displayName} - {taskNumber}", html);
    }

    /// <summary>Goes where the originals had a screenshot of the address book.</summary>
    private static string AddressBlock(string address) =>
        $"""<p style="margin-left:24px"><b>{E(address)}</b></p>""";

    /// <summary>
    /// Shares the shared-mailbox setting deliberately: it is the same knowledge-base article, and
    /// duplicating it would let the two mails drift apart.
    /// </summary>
    private static string MembershipLink(SharedMailboxSettings settings) =>
        string.IsNullOrWhiteSpace(settings.GroupMembershipHelpUrl)
            ? ""
            : $"""<p><a href="{E(settings.GroupMembershipHelpUrl.Trim())}">Editing Group Membership</a></p>""";

    private static string NameList(IReadOnlyList<string> names, string whenEmpty)
    {
        var cleaned = names.Select(n => n.Trim()).Where(n => n.Length > 0).ToArray();
        return cleaned.Length == 0 ? E(whenEmpty) : E(string.Join(", ", cleaned));
    }

    private static string E(string value) => WebUtility.HtmlEncode(value);
}
