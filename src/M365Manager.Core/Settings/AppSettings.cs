using System.Text.Json.Serialization;

namespace M365Manager.Core.Settings;

/// <summary>Root application settings persisted per-user under %APPDATA%\M365Manager.</summary>
public sealed class AppSettings
{
    public SqlSettings Sql { get; set; } = new();
    public M365Settings M365 { get; set; } = new();
    public SmtpSettings Smtp { get; set; } = new();
    public SharedMailboxSettings SharedMailboxes { get; set; } = new();
    public RoomResourceSettings RoomResources { get; set; } = new();
    public TeamsPolicySettings TeamsPolicies { get; set; } = new();
    public TraceSettings Trace { get; set; } = new();
    public NewHireSettings NewHire { get; set; } = new();
}

/// <summary>
/// Settings for the New Hire wizard, which enables a new employee for Teams telephony.
///
/// Its number inventory lives in a separate, long-standing telephony database (historically
/// "uchelper") on the <em>same</em> SQL Server as the app's own database - so only the catalog name
/// is configured here and everything else (server, credentials) is reused from
/// <see cref="SqlSettings"/>.
/// </summary>
public sealed class NewHireSettings
{
    /// <summary>
    /// Catalog name of the telephony database on the server from <see cref="SqlSettings.Server"/>.
    /// Empty disables the New Hire page - it has nothing to read numbers from.
    /// </summary>
    public string Database { get; set; } = DefaultDatabase;

    public const string DefaultDatabase = "uchelper";

    /// <summary>
    /// Also write the SIP address and line URI to the on-premises AD account (the msRTCSIP-*
    /// attributes) before configuring Teams.
    ///
    /// This matters whenever Entra Connect still synchronises those attributes: on-prem is then the
    /// authoritative source, and a number assigned only in Teams is overwritten - or cleared - on
    /// the next sync cycle. Writing the same values on both sides makes the sync a no-op instead.
    /// Turn this off once the SfB schema is no longer synchronised, and the wizard becomes
    /// Teams-only.
    /// </summary>
    public bool WriteAdAttributes { get; set; } = true;

    /// <summary>
    /// Domain controller or domain name to bind to, e.g. "dc01.corp.contoso.com". Empty uses the
    /// machine's own domain (serverless bind), which is what a domain-joined admin PC wants.
    /// The bind always runs as the signed-in Windows user - no credentials are stored.
    /// </summary>
    public string LdapServer { get; set; } = "";

    /// <summary>
    /// Value written to msRTCSIP-DeploymentLocator, which marks the account as homed online.
    /// Legacy: "sipfed.online.lync.com".
    /// </summary>
    public string DeploymentLocator { get; set; } = DefaultDeploymentLocator;

    public const string DefaultDeploymentLocator = "sipfed.online.lync.com";

    /// <summary>
    /// Passed to Set-CsPhoneNumberAssignment. "DirectRouting" for numbers from your own SBC,
    /// "CallingPlan" or "OperatorConnect" when Microsoft supplies them.
    /// </summary>
    public string PhoneNumberType { get; set; } = DefaultPhoneNumberType;

    public const string DefaultPhoneNumberType = "DirectRouting";

    /// <summary>Granted to every new hire. Legacy: "UpgradeToTeams".</summary>
    public string TeamsUpgradePolicyName { get; set; } = DefaultTeamsUpgradePolicyName;

    public const string DefaultTeamsUpgradePolicyName = "UpgradeToTeams";

    /// <summary>
    /// Send the employee a mail once they are enabled, telling them their new number. Legacy did
    /// this unconditionally.
    /// </summary>
    public bool NotifyEmployee { get; set; } = true;

    /// <summary>
    /// Team address notified when a run fails halfway, so somebody picks up the half-configured
    /// account. Legacy hardcoded "LST.GlobalIPTAdmin@ul.com". Empty skips the failure notice.
    /// </summary>
    public string AdminNotificationAddress { get; set; } = "";
}

/// <summary>
/// Entra groups that carry a Teams policy assignment (meeting recording, transcription, Copilot,
/// ...). The tenant has far too many groups to show them all, so the feature only ever looks at
/// groups whose display name starts with one of <see cref="GroupPrefixes"/>.
/// </summary>
public sealed class TeamsPolicySettings
{
    /// <summary>
    /// One prefix per line, e.g. "POL.Teams.MTG.". Each line is its own category: a user holds at
    /// most one group per prefix, and assigning within a prefix replaces whatever they had from
    /// that same prefix. Groups under a different prefix are never touched.
    /// </summary>
    public string GroupPrefixes { get; set; } = DefaultGroupPrefixes;

    public const string DefaultGroupPrefixes = "POL.Teams.MTG.";

    /// <summary>The configured prefixes, trimmed and without blank lines.</summary>
    [JsonIgnore]
    public IReadOnlyList<string> Prefixes =>
        (GroupPrefixes ?? "").Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
}

/// <summary>
/// SMTP relay used for the notification e-mails this app sends (new shared mailbox, owner change,
/// new team, ...). Without it the notification falls back to Graph /me/sendMail, which needs the
/// signed-in admin to have a licensed Exchange mailbox - most admin accounts don't, so the mail
/// silently never arrives. Configuring a relay here is the supported path; it also means the mail
/// comes from a neutral team address instead of whoever happened to run the tool.
/// <see cref="Password"/> is DPAPI-encrypted at rest.
/// </summary>
public sealed class SmtpSettings
{
    /// <summary>Relay host name, e.g. "smtp-relay.contoso.com". Empty disables SMTP entirely.</summary>
    public string Host { get; set; } = "";

    public int Port { get; set; } = 25;

    /// <summary>STARTTLS. Internal relays on port 25 usually don't need it; port 587 does.</summary>
    public bool UseSsl { get; set; }

    /// <summary>Envelope and header sender, e.g. "messaging-team@contoso.com".</summary>
    public string FromAddress { get; set; } = "";

    /// <summary>Friendly name shown next to the sender, e.g. "Enterprise Messaging Services Team".</summary>
    public string FromDisplayName { get; set; } = "";

    /// <summary>Optional Reply-To, when replies should go somewhere other than the sender.</summary>
    public string ReplyToAddress { get; set; } = "";

    /// <summary>Optional; blind-copies every notification (e.g. a team mailbox for the record). Comma- or semicolon-separated.</summary>
    public string BccAddress { get; set; } = "";

    /// <summary>Leave empty for an anonymous/IP-allowlisted relay.</summary>
    public string UserName { get; set; } = "";

    /// <summary>Plaintext in memory; encrypted (DPAPI) when written to disk.</summary>
    public string Password { get; set; } = "";

    /// <summary>A relay is only usable once we know both where to send and who to send as.</summary>
    [JsonIgnore]
    public bool IsConfigured => !string.IsNullOrWhiteSpace(Host) && !string.IsNullOrWhiteSpace(FromAddress);
}

/// <summary>SQL Server connection settings. <see cref="Password"/> is DPAPI-encrypted at rest.</summary>
public sealed class SqlSettings
{
    public string Server { get; set; } = "";
    public string Database { get; set; } = "M365Manager";

    /// <summary>Use Windows/Integrated auth instead of a SQL login.</summary>
    public bool UseIntegratedSecurity { get; set; }

    public string UserId { get; set; } = "";

    /// <summary>Plaintext in memory; encrypted (DPAPI) when written to disk.</summary>
    public string Password { get; set; } = "";
}

/// <summary>M365 / Entra ID app-registration settings for delegated (interactive) sign-in.</summary>
public sealed class M365Settings
{
    public string TenantId { get; set; } = "";
    public string ClientId { get; set; } = "";

    /// <summary>
    /// Default UPN domain (e.g. "global.ul.com" - can differ from the mail domain below). When
    /// someone enters a bare SamAccountName (no "@") to add a group member or resolve an owner,
    /// it's completed as "samAccountName@Domain" before the Graph/Exchange call.
    /// </summary>
    public string Domain { get; set; } = "";

    /// <summary>
    /// Default SMTP/mail domain for newly created mail-enabled objects (e.g. "ul.com" - typically
    /// NOT the same as the UPN domain above). Used to pre-select and default the address-domain
    /// picker when creating a shared mailbox.
    /// </summary>
    public string DefaultMailDomain { get; set; } = "";

    /// <summary>
    /// SharePoint admin center URL (e.g. "https://contoso-admin.sharepoint.com"), used by the
    /// Teams feature to lock down site sharing on internal teams via PnP.PowerShell.
    /// </summary>
    public string SharePointAdminUrl { get; set; } = "";
}

/// <summary>
/// Optional retention/role-assignment policy names applied to new shared mailboxes. Left empty by
/// default and skipped in that case - the legacy scripts hardcoded tenant-specific policy names
/// ("UL Default Role Assignment Policy", "UL MRM Policy - 3 yr Delete") that won't exist in every
/// tenant, so this makes them configurable instead of baked in.
/// </summary>
public sealed class SharedMailboxSettings
{
    public string RetentionPolicyName { get; set; } = "";
    public string RoleAssignmentPolicyName { get; set; } = "";

    /// <summary>
    /// Knowledge-base article on editing group membership, linked from the confirmation mails where
    /// the access groups are explained. Empty omits the link.
    /// </summary>
    public string GroupMembershipHelpUrl { get; set; } = DefaultGroupMembershipHelpUrl;

    /// <summary>
    /// Help links listed at the end of the confirmation mails, one "Label|URL" per line. Empty
    /// omits the block. Taken from the legacy SharedMailboxNew.oft template - the tenant's own
    /// ServiceNow articles, hence configurable rather than baked in.
    /// </summary>
    public string MailboxHelpLinks { get; set; } = DefaultMailboxHelpLinks;

    public const string DefaultGroupMembershipHelpUrl =
        "https://ul.service-now.com/ulsp?id=kb_article&sysparm_article=KB0011504";

    public const string DefaultMailboxHelpLinks = """
        Adding a Shared Mailbox in Outlook|https://ul.service-now.com/ulsp?id=kb_article&sysparm_article=KB0013171
        Sending Mail from a Shared Mailbox|https://ul.service-now.com/ulsp?id=kb_article&sysparm_article=KB0011491
        Creating Signature and Rules in a Shared Mailbox|https://ul.service-now.com/ulsp?id=kb_article&sysparm_article=KB0011356
        """;
}

/// <summary>
/// Settings for room and equipment mailboxes. The legacy scripts hardcoded tenant-specific group
/// names ("MBX.RRS.Owner", "DBS.CRP.RRS.Admins") and flipped the room address domain from ul.com
/// to ul.onmicrosoft.com in March 2025 by editing the script - all of that is configuration here.
/// </summary>
public sealed class RoomResourceSettings
{
    /// <summary>
    /// SMTP domain for room/equipment mailbox addresses - typically the tenant's
    /// &lt;tenant&gt;.onmicrosoft.com domain. Falls back to M365Settings.DefaultMailDomain when empty.
    /// </summary>
    public string RoomMailboxDomain { get; set; } = "";

    /// <summary>
    /// SMTP domain for the room list / delegate / users groups. Legacy kept these on the regular
    /// mail domain even after moving the mailboxes themselves to onmicrosoft.com. Falls back to
    /// M365Settings.DefaultMailDomain when empty.
    /// </summary>
    public string RoomGroupDomain { get; set; } = "";

    /// <summary>Group set as ManagedBy on new room lists and delegate/users groups (legacy: "MBX.RRS.Owner").</summary>
    public string RoomListOwnerGroup { get; set; } = "";

    /// <summary>Global admin group granted FullAccess on every new room (legacy: "DBS.CRP.RRS.Admins").</summary>
    public string GlobalAdminGroup { get; set; } = "";

    /// <summary>Booking window default for the Standard/Custom policy. Legacy: 90.</summary>
    public int DefaultBookingWindowDays { get; set; } = 90;

    /// <summary>Maximum meeting duration default, in minutes. Legacy: 1440 (24h).</summary>
    public int DefaultMaximumDurationMinutes { get; set; } = 1440;
}

/// <summary>
/// Settings for the Trace application's Entra guest accounts. The legacy invite script hardcoded
/// all three of these - including a colleague's name in the signature - so they live here instead.
/// </summary>
public sealed class TraceSettings
{
    /// <summary>Stamped onto every invited guest so the overview can find them again. Legacy: "EXT-CALLIS".</summary>
    public string CompanyName { get; set; } = "EXT-CALLIS";

    /// <summary>Where the invitation link sends the guest after they redeem it.</summary>
    public string RedirectUrl { get; set; } = "https://trace.ul.com";

    /// <summary>
    /// Body of the invitation e-mail. "{RedirectUrl}" and "{DisplayName}" are replaced before
    /// sending; everything else - including the signature - is taken literally.
    /// </summary>
    public string InvitationMessage { get; set; } = DefaultInvitationMessage;

    /// <summary>The bilingual text from the legacy script, with the URL turned into a placeholder.</summary>
    public const string DefaultInvitationMessage = """
        Guten Tag,

        Sie erhalten diese Einladung als Gast zur Registrierung in der UL Solutions Azure Umgebung fuer die Nutzung des Tools Trace. Nach erfolgreicher Registrierung ueber den Link weiter unten, koennen Sie sich anschliessend ueber die URL {RedirectUrl} im Tool anmelden.

        Freundliche Gruesse
        Goekhan Oezdil

        Good day,

        You receive this invitation as a guest to register in the UL Solutions Azure environment for using the Trace tool. After successful registration via the link below, you can then log in to the tool via the URL {RedirectUrl}.

        Best regards
        Gokhan Ozdil
        """;
}
