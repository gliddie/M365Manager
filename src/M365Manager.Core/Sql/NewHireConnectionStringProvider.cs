using M365Manager.Core.Settings;
using M365Manager.Data.NewHire;
using Microsoft.Data.SqlClient;

namespace M365Manager.Core.Sql;

/// <summary>
/// Points the New Hire feature at the legacy telephony database. It sits on the same server and uses
/// the same credentials as the app's own database - only the catalog differs - so everything but
/// <see cref="NewHireSettings.Database"/> is taken from <see cref="SqlSettings"/>. Sibling of
/// <see cref="SettingsConnectionStringProvider"/>.
/// </summary>
public sealed class NewHireConnectionStringProvider : INewHireConnectionStringProvider
{
    private readonly ISettingsService _settings;

    public NewHireConnectionStringProvider(ISettingsService settings)
    {
        _settings = settings;
    }

    public string? GetConnectionString()
    {
        var sql = _settings.Current.Sql;
        var database = _settings.Current.NewHire.Database?.Trim() ?? "";

        if (string.IsNullOrWhiteSpace(sql.Server) || database.Length == 0)
            return null;

        var builder = new SqlConnectionStringBuilder
        {
            DataSource = sql.Server,
            InitialCatalog = database,
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
