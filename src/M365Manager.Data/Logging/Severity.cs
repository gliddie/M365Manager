namespace M365Manager.Data.Logging;

/// <summary>
/// Normalized severity for a log entry. Stored as TINYINT in the database.
/// The legacy 4-char event codes (STAR, INFO, ADD, ERR, FAIL...) map onto these.
/// </summary>
public enum Severity : byte
{
    Debug = 0,
    Info = 1,
    Success = 2,
    Warning = 3,
    Error = 4,
}
