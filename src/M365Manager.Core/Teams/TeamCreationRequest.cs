namespace M365Manager.Core.Teams;

/// <summary>Input for <see cref="ITeamsService.CreateTeamAsync"/> - mirrors the fields the legacy
/// "New Unified Group" WinForms dialog asked for.</summary>
public sealed class TeamCreationRequest
{
    /// <summary>ServiceNow (or equivalent) task number authorizing the creation. Required.</summary>
    public required string TaskNumber { get; init; }

    /// <summary>Location/site code used as the middle segment of the enforced name.</summary>
    public required string Location { get; init; }

    /// <summary>Base team name, before the GRP.&lt;Location&gt;. prefix is applied.</summary>
    public required string Name { get; init; }

    public required string Description { get; init; }

    /// <summary>SamAccountName, UPN or email address of the intended team owner.</summary>
    public required string OwnerIdentity { get; init; }

    public bool IsPublic { get; init; }

    /// <summary>
    /// Internal teams get guest access disabled and their SharePoint site sharing locked down.
    /// External teams keep guest access enabled and are left at the tenant sharing default.
    /// </summary>
    public bool IsInternal { get; init; } = true;
}
