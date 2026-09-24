using System.Globalization;
using System.Text;

namespace M365Manager.Core.Naming;

/// <summary>
/// Turns a human-readable name into something Exchange accepts as an alias and SMTP local part.
/// Display names are never passed through this - they keep whatever the requester wrote.
/// </summary>
public static class ExchangeNaming
{
    /// <summary>
    /// Keeps ASCII letters, digits and dots; transliterates accented Latin letters to their base
    /// letter (Büro -> Buro, José -> Jose) and drops everything else.
    ///
    /// The legacy scripts only ever stripped spaces, "/" and "-" and dealt with anything else -
    /// notably "&amp;" - by hand-adding an entry to KnownAcronyms.csv when it first came up
    /// ("R&amp;I" -> "RI", "C&amp;I" -> "CI"). Anything not yet on that list produced an address
    /// Exchange rejects, since its documented alias character set excludes "&amp;".
    ///
    /// This is deliberately stricter than Exchange itself, which also permits
    /// <c>! # % * + - / = ? ^ _ ~</c>: the naming convention never intends to emit those, and an
    /// address like "UL.Team#1@..." trips up plenty of downstream systems even when Exchange
    /// accepts it.
    /// </summary>
    public static string ToAliasSafe(string? value)
    {
        if (string.IsNullOrEmpty(value))
            return "";

        // "ß" has no canonical decomposition, so expand it before normalizing.
        var expanded = value.Replace("ß", "ss").Replace("ẞ", "SS");
        var normalized = expanded.Normalize(NormalizationForm.FormD);

        var sb = new StringBuilder(normalized.Length);
        foreach (var ch in normalized)
        {
            // FormD splits "ü" into "u" + combining diaeresis; drop the mark, keep the letter.
            if (CharUnicodeInfo.GetUnicodeCategory(ch) == UnicodeCategory.NonSpacingMark)
                continue;

            if (char.IsAsciiLetterOrDigit(ch) || ch == '.')
                sb.Append(ch);
        }

        // Removing characters can leave runs of dots behind. Exchange requires every period to sit
        // between two valid characters, so collapse them and never start or end on one.
        var result = sb.ToString();
        while (result.Contains(".."))
            result = result.Replace("..", ".");

        return result.Trim('.');
    }
}
