using M365Manager.Core.Settings;
using M365Manager.Data.Logging;
using Microsoft.Data.SqlClient;

namespace M365Manager.Core.Sql;

/// <summary>
/// Builds the SQL Server connection string from the user's saved settings.
/// Bridges Core settings to the Data layer's <see cref="IConnectionStringProvider"/>.
/// </summary>
public sealed class SettingsConnectionStringProvider : IConnectionStringProvider
{
    private readonly ISettingsService _settings;

    public SettingsConnectionStringProvider(ISettingsService settings)
    {
        _settings = settings;
    }

    public string? GetConnectionString()
    {
        var sql = _settings.Current.Sql;

        if (string.IsNullOrWhiteSpace(sql.Server) || string.IsNullOrWhiteSpace(sql.Database))
            return null;

        var builder = new SqlConnectionStringBuilder
        {
            DataSource = sql.Server,
            InitialCatalog = sql.Database,
            Encrypt = true,
            TrustServerCertificate = true,
            ConnectTimeout = 10,
            ApplicationName = "M365Manager",
        };

        if (sql.UseIntegratedSecurity)
        {
            builder.IntegratedSecurity = true;
        }
        else
        {
            builder.UserID = sql.UserId;
            builder.Password = sql.Password;
        }

        return builder.ConnectionString;
    }
}
