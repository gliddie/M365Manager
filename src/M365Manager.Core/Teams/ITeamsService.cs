namespace M365Manager.Core.Teams;

/// <summary>
/// Orchestrates creating a Microsoft Team the same way the legacy "New Unified Group / New Team"
/// PowerShell script did: enforced naming, owner licensing check, hide-from-GAL, disabled welcome
/// message, guest-access control, SharePoint sharing lockdown for internal teams, owner
/// notification, and a full step-by-step audit trail in dbo.LogEntries.
/// </summary>
public interface ITeamsService
{
    Task<TeamCreationResult> CreateTeamAsync(TeamCreationRequest request, Action<string>? onProgress = null, CancellationToken ct = default);

    /// <summary>Reads the Teams overview grid from the SQL cache (dbo.Teams/dbo.Groups) - never live from M365.</summary>
    Task<IReadOnlyList<TeamOverviewRow>> GetTeamsOverviewAsync(CancellationToken ct = default);
}
