namespace M365Manager.Data.Naming;

/// <summary>
/// One acronym/term translation applied when building a team name under the
/// GRP.&lt;Location&gt;.&lt;Name&gt; naming convention (see TeamNamingService).
/// Leading/trailing spaces in either field are intentional - they control
/// word-boundary replacement, matching the legacy Check-GroupName behavior.
/// </summary>
public sealed class TeamNameAcronym
{
    public int Id { get; set; }
    public string Acronym { get; set; } = "";
    public string Translation { get; set; } = "";
}
