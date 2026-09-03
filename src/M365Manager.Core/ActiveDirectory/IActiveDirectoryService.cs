namespace M365Manager.Core.ActiveDirectory;

/// <summary>The on-premises AD account behind a SamAccountName.</summary>
/// <param name="DistinguishedName">Used to re-bind for the write, so the search runs only once.</param>
/// <param name="SamAccountName">As stored in AD, which may differ in case from what was typed.</param>
/// <param name="DisplayName">Shown in the wizard and used in the notification mail.</param>
/// <param name="UserPrincipalName">The sign-in name; also what Teams cmdlets are addressed with.</param>
/// <param name="Mail">Primary SMTP address. The SIP address is derived from it, as the legacy tool did.</param>
/// <param name="Office">physicalDeliveryOfficeName - matched against the telephony database's site list.</param>
/// <param name="GivenName">Salutation in the notification mail.</param>
/// <param name="LineUri">Current msRTCSIP-Line, if any. A value here means the user already has a number on-prem.</param>
/// <param name="SipAddress">Current msRTCSIP-PrimaryUserAddress, if any.</param>
public sealed record AdUser(
    string DistinguishedName,
    string SamAccountName,
    string DisplayName,
    string UserPrincipalName,
    string Mail,
    string Office,
    string GivenName,
    string? LineUri,
    string? SipAddress);

/// <summary>Outcome of writing one AD attribute, so the wizard can show which ones landed.</summary>
/// <param name="Attribute">The LDAP attribute name, e.g. "msRTCSIP-Line".</param>
/// <param name="Value">What was written.</param>
/// <param name="Error">Null on success.</param>
public sealed record AdAttributeResult(string Attribute, string Value, string? Error)
{
    public bool Succeeded => Error is null;
}

/// <summary>
/// Reads and writes the on-premises AD account of a new hire.
///
/// Bind runs as the signed-in Windows user over Kerberos - unlike the legacy tool, which carried a
/// service account and its password in the source. That means the admin running the app needs write
/// access to the msRTCSIP-* attributes; if they don't, the write fails loudly here rather than
/// silently succeeding under someone else's rights.
/// </summary>
public interface IActiveDirectoryService
{
    /// <summary>
    /// Finds an account by SamAccountName. Returns null when nothing matches. Throws when the
    /// directory itself is unreachable, which is a different problem and needs a different message.
    /// </summary>
    Task<AdUser?> FindUserAsync(string samAccountName, CancellationToken ct = default);

    /// <summary>
    /// Writes the Skype-for-Business hybrid attributes that Entra Connect synchronises to the cloud:
    /// msRTCSIP-Line, -PrimaryUserAddress, -DeploymentLocator, -FederationEnabled,
    /// -InternetAccessEnabled and -UserEnabled.
    ///
    /// One result per attribute, in the order attempted; a failure on one does not stop the rest,
    /// because a partially configured account still needs the remaining values to be consistent.
    /// </summary>
    Task<IReadOnlyList<AdAttributeResult>> WriteTelephonyAttributesAsync(
        AdUser user,
        string lineUri,
        string sipAddress,
        string deploymentLocator,
        CancellationToken ct = default);
}
