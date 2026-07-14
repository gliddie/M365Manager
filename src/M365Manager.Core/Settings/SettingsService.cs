using System.IO;
using System.Text.Json;
using M365Manager.Core.Security;

namespace M365Manager.Core.Settings;

/// <summary>
/// Loads/saves <see cref="AppSettings"/> as JSON in %APPDATA%\M365Manager\settings.json.
/// Secret fields are encrypted with DPAPI before being written and decrypted on load.
/// </summary>
public sealed class SettingsService : ISettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public string SettingsFilePath { get; }

    public AppSettings Current { get; private set; } = new();

    public SettingsService()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "M365Manager");
        Directory.CreateDirectory(dir);
        SettingsFilePath = Path.Combine(dir, "settings.json");
        Load();
    }

    public AppSettings Load()
    {
        if (!File.Exists(SettingsFilePath))
        {
            Current = new AppSettings();
            return Current;
        }

        try
        {
            var json = File.ReadAllText(SettingsFilePath);
            var settings = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();

            // Decrypt secrets that were encrypted at rest.
            settings.Sql.Password = SecretProtector.Unprotect(settings.Sql.Password) ?? "";

            Current = settings;
        }
        catch
        {
            Current = new AppSettings();
        }

        return Current;
    }

    public void Save(AppSettings settings)
    {
        // Build a copy with encrypted secrets so we never write plaintext to disk.
        var persisted = new AppSettings
        {
            Sql = new SqlSettings
            {
                Server = settings.Sql.Server,
                Database = settings.Sql.Database,
                UseIntegratedSecurity = settings.Sql.UseIntegratedSecurity,
                UserId = settings.Sql.UserId,
                Password = SecretProtector.Protect(settings.Sql.Password) ?? "",
            },
            M365 = new M365Settings
            {
                TenantId = settings.M365.TenantId,
                ClientId = settings.M365.ClientId,
                Domain = settings.M365.Domain,
            },
        };

        var json = JsonSerializer.Serialize(persisted, JsonOptions);
        File.WriteAllText(SettingsFilePath, json);

        // Keep the in-memory copy with plaintext secrets for immediate use.
        Current = settings;
    }
}
