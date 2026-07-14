namespace M365Manager.Core.Settings;

/// <summary>Root application settings persisted per-user under %APPDATA%\M365Manager.</summary>
public sealed class AppSettings
{
    public SqlSettings Sql { get; set; } = new();
    public M365Settings M365 { get; set; } = new();
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
    /// Default UPN domain (e.g. "ul.com"). When someone enters a bare SamAccountName (no "@")
    /// to add a group member, it's completed as "samAccountName@Domain" before the Exchange call.
    /// </summary>
    public string Domain { get; set; } = "";
}
