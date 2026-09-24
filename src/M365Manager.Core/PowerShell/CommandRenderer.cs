using System.Collections;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Text;

namespace M365Manager.Core.PowerShell;

/// <summary>
/// Renders a built pipeline back into the PowerShell syntax that produced it, so the console view
/// shows the actual cmdlet call rather than a description of it. The result is meant to be close
/// enough to paste into a real session.
/// </summary>
public static class CommandRenderer
{
    public static string Render(CommandCollection commands)
    {
        var segments = new List<string>();

        foreach (var command in commands)
        {
            if (command.IsScript)
            {
                // AddScript - show the script itself, collapsed onto one line.
                var script = command.CommandText.Replace("\r\n", "\n").Trim();
                var collapsed = string.Join(" ", script
                    .Split('\n', StringSplitOptions.RemoveEmptyEntries)
                    .Select(l => l.Trim()));
                segments.Add(Truncate(collapsed, 400));
                continue;
            }

            var sb = new StringBuilder(command.CommandText);
            foreach (var parameter in command.Parameters)
            {
                sb.Append(" -").Append(parameter.Name);

                // A switch supplied without a value renders bare, as it would be typed.
                if (parameter.Value is not null)
                    sb.Append(' ').Append(FormatValue(parameter.Value));
            }
            segments.Add(sb.ToString());
        }

        return string.Join(" | ", segments);
    }

    private static string FormatValue(object? value) => value switch
    {
        null => "$null",
        bool b => b ? "$true" : "$false",
        SwitchParameter sw => sw.IsPresent ? "$true" : "$false",
        string s => QuoteIfNeeded(s),
        IDictionary dictionary => FormatDictionary(dictionary),
        IEnumerable sequence and not string => FormatSequence(sequence),
        _ => QuoteIfNeeded(value.ToString() ?? ""),
    };

    private static string FormatDictionary(IDictionary dictionary)
    {
        var entries = new List<string>();
        foreach (DictionaryEntry entry in dictionary)
            entries.Add($"{entry.Key}={FormatValue(entry.Value)}");

        return "@{" + string.Join("; ", entries) + "}";
    }

    private static string FormatSequence(IEnumerable sequence)
    {
        var items = new List<string>();
        foreach (var item in sequence)
        {
            // Long member lists would swamp the line; say how many were left out instead.
            if (items.Count == 10)
            {
                items.Add($"... ({sequence.Cast<object?>().Count()} total)");
                break;
            }
            items.Add(FormatValue(item));
        }

        return items.Count == 0 ? "@()" : string.Join(",", items);
    }

    private static string QuoteIfNeeded(string value)
    {
        if (value.Length == 0)
            return "\"\"";

        // Bare words are fine for anything without whitespace or PowerShell-significant characters.
        var needsQuotes = value.Any(c => char.IsWhiteSpace(c) || "\"'$`;,|&(){}[]".Contains(c));
        return needsQuotes ? "\"" + value.Replace("\"", "`\"") + "\"" : value;
    }

    private static string Truncate(string value, int max)
        => value.Length <= max ? value : value[..max] + " ...";
}
