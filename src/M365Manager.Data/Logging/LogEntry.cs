namespace M365Manager.Data.Logging;

/// <summary>
/// One row in dbo.LogEntries. Replaces a line from the legacy text log files.
/// </summary>
public sealed class LogEntry
{
    public long Id { get; set; }

    /// <summary>When the event happened (UTC).</summary>
    public DateTime TimestampUtc { get; set; }

    /// <summary>Signed-in M365 admin performing the change (their UPN).</summary>
    public string? UserUpn { get; set; }

    /// <summary>Windows account running the app (whoami).</summary>
    public string? WindowsUser { get; set; }

    /// <summary>Computer the action was launched from.</summary>
    public string? MachineName { get; set; }

    /// <summary>Functional area, e.g. "UserMailbox", "Licensing", "Groups".</summary>
    public string? Area { get; set; }

    /// <summary>Operation / former script name, e.g. "AddFolderPermissions".</summary>
    public string? Action { get; set; }

    /// <summary>Object acted on (mailbox / UPN / group / domain).</summary>
    public string? TargetObject { get; set; }

    /// <summary>Legacy 4-char event code (STAR, INFO, ADD, REMO, ERR, FAIL...).</summary>
    public string? EventCode { get; set; }

    /// <summary>Normalized severity.</summary>
    public Severity Severity { get; set; } = Severity.Info;

    public string? Message { get; set; }

    /// <summary>Groups all entries belonging to one operation/session.</summary>
    public Guid? CorrelationId { get; set; }

    /// <summary>Row insertion time (UTC).</summary>
    public DateTime CreatedAtUtc { get; set; }
}
