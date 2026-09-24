namespace M365Manager.Core.SharedMailboxes;

public sealed class ChangeSharedMailboxOwnerResult
{
    public required bool Succeeded { get; init; }

    /// <summary>Comma-separated display names of the resulting owners (after the change), formatted "Owner: First Last, ...".</summary>
    public string? UpdatedOwners { get; init; }

    public string? ErrorMessage { get; init; }

    /// <summary>Set when the change went through but the owner notification e-mail did not. See <see cref="SharedMailboxCreationResult.WarningMessage"/>.</summary>
    public string? WarningMessage { get; init; }

    public Guid CorrelationId { get; init; }

    public static ChangeSharedMailboxOwnerResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}
