namespace M365Manager.Data.NewHire;

/// <summary>
/// Supplies the connection string for the legacy telephony database. Separate from
/// <see cref="M365Manager.Data.Logging.IConnectionStringProvider"/> because it points at a different
/// catalog on the same server - the app's own database holds the audit log, this one holds the
/// number inventory.
/// </summary>
public interface INewHireConnectionStringProvider
{
    /// <summary>Returns the connection string, or null when the telephony database is not configured.</summary>
    string? GetConnectionString();
}

/// <summary>
/// A free number found inside a DID range, ready to be offered to the operator.
/// </summary>
/// <param name="Did">The bare digits, e.g. "4969120051042" - no leading "+".</param>
public sealed record FreeDid(string Did)
{
    /// <summary>E.164 form, which is what Teams wants on the wire.</summary>
    public string E164 => "+" + Did;

    public override string ToString() => E164;
}

/// <summary>
/// What the telephony database knows about an employee before the wizard touches anything.
///
/// Licensing deliberately does not appear here: it is read live from Entra instead, because the
/// <c>e3licensed</c> table only models E3 and this tenant has moved to E5 - see
/// <c>NewHireService.GetLicenseStateAsync</c>.
/// </summary>
public sealed class UcEmployeeState
{
    /// <summary>The endpoints row, when the user already holds a number. Null means not yet configured.</summary>
    public UcEndpoint? Endpoint { get; init; }

    /// <summary>True when the endpoints row says EnterpriseVoiceEnabled - i.e. this is not a new hire.</summary>
    public bool IsAlreadyEnabled =>
        string.Equals(Endpoint?.EnterpriseVoiceEnabled, "true", StringComparison.OrdinalIgnoreCase);
}

/// <summary>
/// Everything the New Hire wizard reads from and writes to the legacy telephony database.
/// </summary>
public interface INewHireRepository
{
    /// <summary>True once a server and a telephony database name are configured.</summary>
    bool IsConfigured { get; }

    /// <summary>Opens a connection and runs a trivial query, so Settings can verify the catalog name.</summary>
    Task<bool> TestConnectionAsync(CancellationToken ct = default);

    /// <summary>Every site that has a location code, ordered by name - the office picker's list.</summary>
    Task<IReadOnlyList<UcLocation>> GetLocationsAsync(CancellationToken ct = default);

    /// <summary>Looks a site up by its AD office name. Case-insensitive via the column's collation.</summary>
    Task<UcLocation?> FindLocationByNameAsync(string officeName, CancellationToken ct = default);

    /// <summary>Endpoint and licence state for one employee.</summary>
    Task<UcEmployeeState> GetEmployeeStateAsync(string samAccountName, CancellationToken ct = default);

    /// <summary>The usable (SDAP = 1) number blocks for a site, least-used first.</summary>
    Task<IReadOnlyList<UcDidRange>> GetDidRangesAsync(string locationCode, CancellationToken ct = default);

    /// <summary>
    /// Enumerates <paramref name="range"/> and drops every number already assigned to an endpoint or
    /// explicitly blocked.
    /// </summary>
    Task<IReadOnlyList<FreeDid>> FindFreeDidsAsync(UcDidRange range, CancellationToken ct = default);

    /// <summary>
    /// Opens a run in "newhirelog" before anything is changed and returns its id, so a run that dies
    /// halfway still leaves a trace. Every step outcome is written back against this id.
    /// </summary>
    Task<int> StartRunAsync(string samAccountName, string adminAccount, string ticket, CancellationToken ct = default);

    /// <summary>Writes the per-step outcome of a finished run back onto the row opened by <see cref="StartRunAsync"/>.</summary>
    Task CompleteRunAsync(int newHireLogId, Action<UcNewHireLog> apply, CancellationToken ct = default);

    /// <summary>
    /// Records the newly configured endpoint so the number stops showing up as free. The nightly
    /// import would eventually do this too, but not before the next operator has already handed the
    /// same number to somebody else.
    /// </summary>
    Task AddEndpointAsync(UcEndpoint endpoint, CancellationToken ct = default);
}
