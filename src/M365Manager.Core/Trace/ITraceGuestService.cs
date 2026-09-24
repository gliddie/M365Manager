namespace M365Manager.Core.Trace;

/// <summary>
/// Entra guest accounts for the Trace application, replacing the standalone
/// New-MgInvitation script: invite the customer, stamp the configured company name onto the new
/// guest so it can be found again, and audit both steps in dbo.LogEntries.
/// </summary>
public interface ITraceGuestService
{
    Task<TraceGuestInvitationResult> InviteAsync(TraceGuestInvitationRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>
    /// Every guest carrying the configured company name, read live from Entra. There is no SQL
    /// cache here - the list is small and always current, unlike the mailbox/team overviews.
    /// </summary>
    Task<IReadOnlyList<TraceGuestRow>> GetGuestsAsync(CancellationToken ct = default);
}
