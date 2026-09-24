namespace M365Manager.Data.Logging;

/// <summary>
/// Filter for <see cref="ILogService.QueryAsync"/>. Every property is optional; the ones that are
/// set are combined with AND. Filtering happens in SQL, not on an already-loaded page, so a date
/// range reaches entries far older than the default "most recent" view.
/// </summary>
public sealed class LogQuery
{
    /// <summary>Inclusive lower bound on <see cref="LogEntry.TimestampUtc"/>.</summary>
    public DateTime? FromUtc { get; init; }

    /// <summary>Inclusive upper bound on <see cref="LogEntry.TimestampUtc"/>.</summary>
    public DateTime? ToUtc { get; init; }

    /// <summary>Exact match on the signed-in admin's UPN.</summary>
    public string? UserUpn { get; init; }

    /// <summary>Exact match on the functional area, e.g. "SharedMailboxes".</summary>
    public string? Area { get; init; }

    /// <summary>Exact match on the operation, e.g. "CreateSharedMailbox".</summary>
    public string? Action { get; init; }

    /// <summary>Substring match on the ticket/task number.</summary>
    public string? TaskNumber { get; init; }

    /// <summary>Exact severity. Null means every severity.</summary>
    public Severity? Severity { get; init; }

    /// <summary>Substring match across message, target, action and event code.</summary>
    public string? SearchText { get; init; }

    /// <summary>Safety valve so a wide-open filter can't pull the whole table into the grid.</summary>
    public int MaxResults { get; init; } = 1000;
}

/// <summary>Distinct values behind the log filter's drop-downs, read from the whole table.</summary>
public sealed class LogFilterOptions
{
    public IReadOnlyList<string> Users { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> Areas { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> Actions { get; init; } = Array.Empty<string>();
}
