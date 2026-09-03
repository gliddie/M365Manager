namespace M365Manager.Core.SharedMailboxes;

public sealed class SharedMailboxCreationResult
{
    public required bool Succeeded { get; init; }
    public string? DisplayName { get; init; }
    public string? PrimarySmtpAddress { get; init; }
    public string? ErrorMessage { get; init; }

    /// <summary>
    /// Set when the mailbox was created but a best-effort follow-up step failed - currently only
    /// the owner confirmation e-mail. <see cref="Succeeded"/> stays true; the UI shows this next to
    /// the success message so a broken SMTP relay doesn't go unnoticed.
    /// </summary>
    public string? WarningMessage { get; init; }

    public Guid CorrelationId { get; init; }

    public static SharedMailboxCreationResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}
