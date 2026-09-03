namespace M365Manager.Data.Naming;

/// <summary>
/// Applies the acronym/term translations from dbo.TeamNameAcronyms to a name.
/// </summary>
public static class AcronymSubstitution
{
    /// <summary>
    /// Replaces every acronym with its translation, <b>case-insensitively</b>.
    ///
    /// Case-insensitivity is not incidental: the legacy scripts did this with PowerShell's
    /// -Replace operator, which is case-insensitive by default, on a name that had already been
    /// run through ToTitleCase. So "HVAC" in the acronym table still matched the title-cased
    /// "Hvac". A plain C# string.Replace is case-sensitive and would silently match nothing,
    /// leaving the whole acronym table inert.
    /// </summary>
    public static string Apply(string text, IEnumerable<TeamNameAcronym> acronyms)
    {
        var result = text ?? "";

        foreach (var acronym in acronyms)
        {
            if (!string.IsNullOrEmpty(acronym.Acronym))
                result = ReplaceIgnoreCase(result, acronym.Acronym, acronym.Translation ?? "");
        }

        return result;
    }

    private static string ReplaceIgnoreCase(string haystack, string needle, string replacement)
    {
        var index = haystack.IndexOf(needle, StringComparison.OrdinalIgnoreCase);
        if (index < 0)
            return haystack;

        var sb = new System.Text.StringBuilder(haystack.Length);
        var position = 0;

        while (index >= 0)
        {
            sb.Append(haystack, position, index - position).Append(replacement);
            position = index + needle.Length;
            index = haystack.IndexOf(needle, position, StringComparison.OrdinalIgnoreCase);
        }

        return sb.Append(haystack, position, haystack.Length - position).ToString();
    }
}
