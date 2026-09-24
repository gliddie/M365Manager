using System.Collections;
using System.Management.Automation;

namespace M365Manager.Core.PowerShell;

/// <summary>Reading values out of the <see cref="PSObject"/>s the Exchange cmdlets return.</summary>
public static class PsValues
{
    /// <summary>
    /// A multi-valued Exchange property (ManagedBy, ResourceDelegates, EmailAddresses, ...) as a
    /// list of strings.
    ///
    /// Reading such a property off a PSObject is not reliable across cmdlet flavours: a real run
    /// against Exchange Online produced a ManagedBy whose three owners arrived already flattened
    /// into the single string "&lt;guid&gt; Chiara Clemente Lori Hemingway", which then got used as
    /// one identity. Splitting that back apart is impossible - the names contain spaces themselves.
    /// So where each element has to be usable on its own, ask PowerShell to expand the property
    /// instead (see <see cref="ExpandProperty"/>); this method is for the cases that only need a
    /// display value, and it at least unwraps a PSObject rather than stringifying the wrapper.
    /// </summary>
    public static List<string> ToStringList(PSObject? source, string propertyName)
    {
        var list = new List<string>();
        var value = source?.Properties[propertyName]?.Value;

        if (value is PSObject wrapper)
            value = wrapper.BaseObject;

        switch (value)
        {
            case null:
                break;

            // Checked before IEnumerable - a string is a sequence of chars.
            case string text:
                Add(list, text);
                break;

            case IEnumerable sequence:
                foreach (var item in sequence)
                    Add(list, (item is PSObject element ? element.BaseObject : item)?.ToString());
                break;

            default:
                Add(list, value.ToString());
                break;
        }

        return list;
    }

    /// <summary>
    /// Reads a multi-valued property by piping the cmdlet's output through
    /// <c>Select-Object -ExpandProperty</c>, so every element arrives as its own pipeline object no
    /// matter how the property itself is shaped. Use this whenever the individual values are going
    /// to be used as identities, not merely displayed.
    /// </summary>
    /// <param name="host">The shared PowerShell host.</param>
    /// <param name="command">Cmdlet producing the object, e.g. "Get-DistributionGroup".</param>
    /// <param name="identity">Value for that cmdlet's -Identity parameter.</param>
    /// <param name="propertyName">The multi-valued property to expand, e.g. "ManagedBy".</param>
    public static async Task<List<string>> ExpandProperty(
        PowerShellHost host, string command, string identity, string propertyName, CancellationToken ct = default)
    {
        var items = await host.InvokeAsync(ps => ps
            .AddCommand(command)
            .AddParameter("Identity", identity)
            .AddParameter("ErrorAction", "Stop")
            .AddCommand("Select-Object")
            .AddParameter("ExpandProperty", propertyName), ct: ct);

        var list = new List<string>();
        foreach (var item in items)
            Add(list, (item?.BaseObject ?? item)?.ToString());

        return list;
    }

    private static void Add(List<string> list, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
            list.Add(value.Trim());
    }
}
