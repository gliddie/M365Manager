namespace M365Manager.Data.Logging;

/// <summary>
/// Central logging service. Every admin action writes here instead of to text files.
/// </summary>
public interface ILogService
{
    /// <summary>Writes one log entry. Fills in machine/timestamp defaults if not set.</summary>
    Task WriteAsync(LogEntry entry, CancellationToken ct = default);

    /// <summary>Returns the most recent entries, newest first.</summary>
    Task<IReadOnlyList<LogEntry>> GetRecentAsync(int count = 200, CancellationToken ct = default);

    /// <summary>Returns the entries matching <paramref name="query"/>, newest first. Filtering runs in SQL.</summary>
    Task<IReadOnlyList<LogEntry>> QueryAsync(LogQuery query, CancellationToken ct = default);

    /// <summary>Distinct users/areas/actions across the whole log, for the filter drop-downs.</summary>
    Task<LogFilterOptions> GetFilterOptionsAsync(CancellationToken ct = default);

    /// <summary>Checks whether the configured SQL database is reachable.</summary>
    Task<bool> TestConnectionAsync(CancellationToken ct = default);
}
