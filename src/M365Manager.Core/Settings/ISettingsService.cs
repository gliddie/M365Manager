namespace M365Manager.Core.Settings;

public interface ISettingsService
{
    /// <summary>The current in-memory settings (secrets decrypted).</summary>
    AppSettings Current { get; }

    /// <summary>Reloads settings from disk.</summary>
    AppSettings Load();

    /// <summary>Persists settings to disk (secrets encrypted with DPAPI).</summary>
    void Save(AppSettings settings);

    /// <summary>Full path to the settings file (for display / troubleshooting).</summary>
    string SettingsFilePath { get; }
}
