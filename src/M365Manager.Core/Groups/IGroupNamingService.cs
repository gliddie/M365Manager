namespace M365Manager.Core.Groups;

/// <summary>The names derived from the operator's dot-notation input.</summary>
public sealed record GroupNames(string DisplayName, string Alias, string Address);

/// <summary>
/// Builds distribution/security group names from the dot-notation convention
/// (LST.EMEA.Some Team Name), reproducing NewDLForm.ps1 - minus its ToTitleCase step.
/// </summary>
public interface IGroupNamingService
{
    /// <param name="rawName">Dot-notation name as typed, e.g. "LST.EMEA.Some Team".</param>
    /// <param name="isTemporary">Inserts "TMP" as the second segment, shifting the rest.</param>
    /// <param name="removeCharacter">Optional single character stripped out of the third segment (legacy "Remove Special Character").</param>
    /// <param name="domain">SMTP domain for the address.</param>
    Task<GroupNames> BuildAsync(
        string rawName,
        bool isTemporary,
        string removeCharacter,
        string domain,
        CancellationToken ct = default);
}
