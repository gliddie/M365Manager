namespace M365Manager.Core.Teams;

public sealed class TeamCreationResult
{
    public required bool Succeeded { get; init; }
    public string? DisplayName { get; init; }
    public string? PrimarySmtpAddress { get; init; }
    public string? GroupId { get; init; }
    public string? ResolvedOwnerUpn { get; init; }
    public string? ErrorMessage { get; init; }
    public Guid CorrelationId { get; init; }

    public static TeamCreationResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}
