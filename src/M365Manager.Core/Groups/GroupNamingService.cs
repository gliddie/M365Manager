using System.Text;
using System.Text.RegularExpressions;
using M365Manager.Core.Naming;
using M365Manager.Data.Naming;

namespace M365Manager.Core.Groups;

/// <summary>
/// See <see cref="IGroupNamingService"/>. Follows NewDLForm.ps1's segment rules, with one
/// deliberate omission: the legacy `ToTitleCase($name.ToLower())` is NOT applied. That step is what
/// forced 103 of the 117 KnownAcronyms rows to be identity mappings purely to repair the casing it
/// destroyed ("IT" -> "It"); names are used as typed instead.
/// </summary>
public sealed partial class GroupNamingService : IGroupNamingService
{
    private readonly ITeamNamingRepository _repository;

    public GroupNamingService(ITeamNamingRepository repository)
    {
        _repository = repository;
    }

    public async Task<GroupNames> BuildAsync(
        string rawName,
        bool isTemporary,
        string removeCharacter,
        string domain,
        CancellationToken ct = default)
    {
        var name = ApplySpecialCases((rawName ?? "").Trim());
        var segments = name.Split('.');

        if (segments.Length == 0 || segments[0].Length == 0)
            return new GroupNames("", "", "");

        var acronyms = await _repository.GetAllAsync(ct);
        var result = (string[])segments.Clone();

        for (var i = 0; i < segments.Length; i++)
        {
            if (i == 0)
            {
                // Prefix (LST / DSG / DST) is always upper-case.
                result[0] = segments[0].ToUpperInvariant();
            }
            else if (i == 1)
            {
                if (isTemporary && !segments[1].Equals("TMP", StringComparison.OrdinalIgnoreCase))
                {
                    // The location moves into the third segment and "TMP" takes its place.
                    var location = segments[1].Length == 3 ? segments[1].ToUpperInvariant() : segments[1];
                    if (result.Length > 2)
                        result[2] = $"{location} {result[2]}";
                    result[1] = "TMP";
                }
                else
                {
                    result[1] = NormalizeLocation(segments[1]);
                }
            }
            else
            {
                // With TMP in play, segment 2 already holds the merged "<location> <name>" built
                // above and must not be overwritten with the original segment.
                if (!isTemporary || i > 2)
                    result[i] = segments[i];

                result[i] = AcronymSubstitution.Apply(result[i], acronyms);
            }
        }

        var displayName = Reassemble(result, removeCharacter);

        // Legacy: $DLName -Replace '[ &/,$#_-]','' - note the "&" is stripped here, unlike the
        // shared-mailbox scripts. ToAliasSafe then removes anything else Exchange won't accept.
        var localPart = ExchangeNaming.ToAliasSafe(AddressStripRegex().Replace(displayName, ""));

        return new GroupNames(
            DisplayName: displayName,
            Alias: localPart,
            Address: localPart.Length == 0 ? "" : $"{localPart}@{(domain ?? "").Trim()}");
    }

    /// <summary>Segments 0-2 are joined with dots, everything after that with spaces.</summary>
    private static string Reassemble(string[] segments, string removeCharacter)
    {
        var sb = new StringBuilder();

        for (var i = 0; i < segments.Length; i++)
        {
            var value = segments[i];

            // The legacy "Remove Special Character" box only ever applied to the third segment.
            if (i == 2 && !string.IsNullOrEmpty(removeCharacter))
                value = value.Replace(removeCharacter, " ");

            if (i == 0)
                sb.Append(value);
            else if (i <= 2)
                sb.Append('.').Append(value);
            else
                sb.Append(' ').Append(value);
        }

        return sb.ToString().Trim();
    }

    /// <summary>Second segment: upper-cased for short codes and EMEA*, otherwise left as typed.</summary>
    private static string NormalizeLocation(string segment)
    {
        if (segment.StartsWith("EMEA", StringComparison.OrdinalIgnoreCase))
            return segment.ToUpperInvariant();

        if (segment.Contains("CommercialOperations", StringComparison.OrdinalIgnoreCase))
            return "CommercialOperations";

        return segment.Length <= 3 ? segment.ToUpperInvariant() : segment;
    }

    /// <summary>
    /// Three names carry an embedded dot that must NOT be treated as a segment separator; legacy
    /// rewrote them before splitting (NewDLForm.ps1:256-270).
    /// </summary>
    private static string ApplySpecialCases(string name)
    {
        if (name.Contains("UL.RS.", StringComparison.OrdinalIgnoreCase))
            return ReplaceIgnoreCase(name, "UL.RS.", "UL.RS ");

        if (name.Contains("UL.IMS.", StringComparison.OrdinalIgnoreCase))
            return ReplaceIgnoreCase(name, "UL.IMS.", "UL.IMS ");

        if (name.Contains("LST.COMMERCIAL.OPERATIONS", StringComparison.OrdinalIgnoreCase))
            return ReplaceIgnoreCase(name, "LST.COMMERCIAL.OPERATIONS", "LST.CommercialOperations");

        return name;
    }

    private static string ReplaceIgnoreCase(string haystack, string needle, string replacement)
        => Regex.Replace(haystack, Regex.Escape(needle), replacement.Replace("$", "$$"), RegexOptions.IgnoreCase);

    // Legacy address rule, verbatim: '[ &/,$#_-]'
    [GeneratedRegex(@"[ &/,$#_\-]")]
    private static partial Regex AddressStripRegex();
}
