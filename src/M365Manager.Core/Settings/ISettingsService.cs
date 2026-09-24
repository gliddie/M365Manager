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

    /// <summary>
    /// Why the last <see cref="Load"/> could not read the file, or part of it - null when it read
    /// cleanly. Anything non-null means some section fell back to defaults, so saving now would
    /// write those defaults over whatever is still in the file. Surface it; do not swallow it.
    /// </summary>
    string? LastLoadError { get; }

    /// <summary>
    /// Where the unreadable settings file was copied to before the app carried on with defaults,
    /// so the user can recover values by hand. Null when nothing was quarantined.
    /// </summary>
    string? QuarantinedFilePath { get; }
}
