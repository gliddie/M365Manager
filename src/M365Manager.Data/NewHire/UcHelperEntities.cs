namespace M365Manager.Data.NewHire;

/// <summary>
/// A site/office in the legacy telephony database (table "locationconfiguration"). The AD office
/// name on a user's account is matched against <see cref="Name"/>; everything the wizard needs to
/// configure them - which DID ranges to draw from and which Teams policies to grant - hangs off
/// <see cref="LocationCode"/>.
/// </summary>
public sealed class UcLocation
{
    public int Id { get; set; }

    /// <summary>Office name exactly as it appears in the AD "physicalDeliveryOfficeName" attribute.</summary>
    public string Name { get; set; } = "";

    /// <summary>
    /// Short code for the site, e.g. "DEFRK". Doubles as the name of all four Teams policies granted
    /// to a new hire (voice routing, dial plan, emergency calling, emergency call routing) - that is
    /// how the legacy tool worked and how the policies are named in the tenant.
    /// </summary>
    public string LocationCode { get; set; } = "";

    public string? City { get; set; }
    public string? Country { get; set; }

    /// <summary>1 = the site is live on Teams voice, 2 = explicitly not enabled, null/0 = never configured.</summary>
    public byte? TeamsVoiceEnabled { get; set; }
}

/// <summary>
/// One contiguous block of phone numbers assigned to a site (table "did"). Numbers are stored as
/// digit strings without a leading "+".
/// </summary>
public sealed class UcDidRange
{
    public int Id { get; set; }

    /// <summary>First number of the block, e.g. "4969120051000".</summary>
    public string DidStart { get; set; } = "";

    /// <summary>Last number of the block, inclusive.</summary>
    public string DidEnd { get; set; } = "";

    public string LocationCode { get; set; } = "";

    /// <summary>Percentage of the block already handed out; shown so the operator can spread usage.</summary>
    public int? Utilization { get; set; }

    public string? Notes { get; set; }

    /// <summary>1 = the block may be used for user DIDs. Blocks with 0 are trunks, service numbers, ...</summary>
    public byte Sdap { get; set; }

    public string? Provider { get; set; }
}

/// <summary>
/// A configured telephony endpoint (table "endpoints") - the inventory of who already holds which
/// number. Kept up to date by the nightly import; the wizard reads it to tell whether an employee is
/// already enabled and writes a row once it has enabled one.
/// </summary>
public sealed class UcEndpoint
{
    public int Id { get; set; }

    public string? DisplayName { get; set; }
    public string? SamAccountName { get; set; }
    public string? SipAddress { get; set; }
    public string? LineUri { get; set; }

    /// <summary>The bare number, no "tel:+" prefix - this is what the free-number search excludes.</summary>
    public string? Did { get; set; }

    public string? Type { get; set; }
    public string? VoicePolicy { get; set; }
    public string? DialPlan { get; set; }

    /// <summary>Stored as the strings "true"/"false" by the legacy importer, not as a bit.</summary>
    public string? EnterpriseVoiceEnabled { get; set; }

    public string? HostedVoiceMail { get; set; }
    public string? EmergencyCallingPolicy { get; set; }
    public string? EmergencyCallRoutingPolicy { get; set; }
    public string? Office { get; set; }
}

/// <summary>
/// A number that must never be handed out even though it falls inside a usable range (table
/// "blockeddids") - reserved, ported away, or burnt by a previous assignment.
/// </summary>
public sealed class UcBlockedDid
{
    public int Id { get; set; }
    public string Did { get; set; } = "";
    public string? Notes { get; set; }
}

/// <summary>
/// Licence state per user (table "e3licensed"), refreshed by a nightly import. A new hire without an
/// E3 licence cannot be enabled for voice yet - the licence has to land first.
/// </summary>
public sealed class UcE3License
{
    public int Id { get; set; }
    public string SamAccountName { get; set; } = "";

    /// <summary>"1" when licensed. A string in the legacy schema, not a bit.</summary>
    public string? E3Licensed { get; set; }

    /// <summary>"1" when the user additionally holds a Teams Phone licence.</summary>
    public string? HasPhoneLicense { get; set; }
}

/// <summary>
/// One run of the wizard (table "newhirelog"). The row is inserted before any change is made and
/// updated with the per-step outcome afterwards, so a run that dies halfway still leaves a trace.
///
/// Every step is an int - 1 succeeded, 0 failed - with a matching *Error column, which is how the
/// legacy PowerShell script recorded its results. That shape is kept so the existing reports over
/// this table keep working.
/// </summary>
public sealed class UcNewHireLog
{
    public int NewHireLogId { get; set; }

    public string? SamAccountName { get; set; }

    /// <summary>Who ran the wizard.</summary>
    public string AdminAccount { get; set; } = "not set";

    public DateTime Date { get; set; }

    /// <summary>ServiceNow ticket authorizing the change.</summary>
    public string? SnTicket { get; set; }

    public int AllVariablesSet { get; set; }
    public int CrdFileExisting { get; set; }
    public int CrdFileUpToDate { get; set; }
    public int TeamsConnected { get; set; }
    public string TeamsConnectedError { get; set; } = "not set";
    public int MsolConnected { get; set; }
    public string MsolConnectedError { get; set; } = "not set";

    public int MsRtcSipDeploymentLocator { get; set; }
    public string MsRtcSipDeploymentLocatorError { get; set; } = "not set";
    public int MsRtcSipFederationEnabled { get; set; }
    public string MsRtcSipFederationEnabledError { get; set; } = "not set";
    public int MsRtcSipInternetAccessEnabled { get; set; }
    public string MsRtcSipInternetAccessEnabledError { get; set; } = "not set";
    public int MsRtcSipUserEnabled { get; set; }
    public string MsRtcSipUserEnabledError { get; set; } = "not set";
    public int MsRtcSipLine { get; set; }
    public string MsRtcSipLineError { get; set; } = "not set";
    public int MsRtcSipPrimaryUserAddress { get; set; }
    public string MsRtcSipPrimaryUserAddressError { get; set; } = "not set";

    public int VoiceRoutingPolicy { get; set; }
    public string VoiceRoutingPolicyError { get; set; } = "not set";
    public int DialPlan { get; set; }
    public string DialPlanError { get; set; } = "not set";
    public int EmergencyCallingPolicy { get; set; }
    public string EmergencyCallingPolicyError { get; set; } = "not set";
    public int EmergencyCallRoutingPolicy { get; set; }
    public string EmergencyCallRoutingPolicyError { get; set; } = "not set";
    public int TeamsUpgradePolicy { get; set; }
    public string TeamsUpgradePolicyError { get; set; } = "not set";

    public int EvEnabled { get; set; }
    public string EvEnabledError { get; set; } = "not set";

    public int SqlServerModule { get; set; }
    public int PhoneLicense { get; set; }
    public string PhoneLicenseError { get; set; } = "not set";
}
