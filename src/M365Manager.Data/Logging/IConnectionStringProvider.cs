namespace M365Manager.Data.Logging;

/// <summary>
/// Supplies the SQL Server connection string at runtime.
/// Implemented in the Core layer from the user's saved settings, so the
/// Data layer stays free of any settings/DPAPI concerns.
/// </summary>
public interface IConnectionStringProvider
{
    /// <summary>Returns the connection string, or null if SQL is not yet configured.</summary>
    string? GetConnectionString();
}
