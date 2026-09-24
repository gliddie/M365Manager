using System.IO;
using System.Text.RegularExpressions;
using M365Manager.Data.Rooms;

namespace M365Manager.ViewModels;

/// <summary>
/// Reads the legacy E:\O365AdminShared\Data\RoomTimeZones.csv into <see cref="RoomSite"/> rows.
///
/// That file had three columns - Code, Zone, Admins - which RoomResourceNewForm.ps1 read via
/// Import-Csv. It had no Region column; the region was picked from a separate list box and only
/// survives inside the admin group name ("MBX.EU.RRS.Admins"), so it's recovered from there.
/// </summary>
public static partial class RoomSiteCsv
{
    public static IReadOnlyList<RoomSite> Parse(IReadOnlyList<string> lines)
    {
        var sites = new List<RoomSite>();
        if (lines.Count == 0)
            return sites;

        var header = SplitCsvLine(lines[0]);
        var codeIndex = IndexOf(header, "Code", "SiteCode", "Site");
        var zoneIndex = IndexOf(header, "Zone", "TimeZone", "Time zone");
        var adminIndex = IndexOf(header, "Admins", "RegionalAdminGroup", "Admin group");
        var regionIndex = IndexOf(header, "Region");

        // No recognizable header: treat the file as headerless Code,Zone,Admins.
        var startRow = 1;
        if (codeIndex < 0 || zoneIndex < 0)
        {
            codeIndex = 0;
            zoneIndex = 1;
            adminIndex = header.Count > 2 ? 2 : -1;
            regionIndex = -1;
            startRow = 0;
        }

        for (var i = startRow; i < lines.Count; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i]))
                continue;

            var fields = SplitCsvLine(lines[i]);
            var code = Field(fields, codeIndex);
            var zone = Field(fields, zoneIndex);
            if (code.Length == 0 || zone.Length == 0)
                continue;

            var admins = Field(fields, adminIndex);
            var region = Field(fields, regionIndex);
            if (region.Length == 0)
                region = RegionFromAdminGroup(admins);

            sites.Add(new RoomSite
            {
                SiteCode = code.ToUpperInvariant(),
                TimeZone = zone,
                Region = region,
                RegionalAdminGroup = admins,
            });
        }

        return sites;
    }

    private static string Field(IReadOnlyList<string> fields, int index)
        => index >= 0 && index < fields.Count ? fields[index].Trim() : "";

    private static int IndexOf(IReadOnlyList<string> header, params string[] names)
    {
        for (var i = 0; i < header.Count; i++)
        {
            var value = header[i].Trim();
            if (names.Any(n => string.Equals(value, n, StringComparison.OrdinalIgnoreCase)))
                return i;
        }
        return -1;
    }

    /// <summary>"MBX.EU.RRS.Admins" -> "EU". Empty when the group doesn't follow that shape.</summary>
    private static string RegionFromAdminGroup(string adminGroup)
    {
        var match = AdminGroupRegex().Match(adminGroup ?? "");
        return match.Success ? match.Groups[1].Value.ToUpperInvariant() : "";
    }

    [GeneratedRegex(@"^MBX\.([A-Za-z]{2,3})\.RRS\.Admins", RegexOptions.IgnoreCase)]
    private static partial Regex AdminGroupRegex();

    /// <summary>Minimal RFC 4180 splitter - the legacy file quoted the Zone column ("W. Europe Standard Time").</summary>
    private static List<string> SplitCsvLine(string line)
    {
        var fields = new List<string>();
        var current = new System.Text.StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];

            if (inQuotes)
            {
                if (ch == '"')
                {
                    // A doubled quote inside a quoted field is a literal quote.
                    if (i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    current.Append(ch);
                }
                continue;
            }

            switch (ch)
            {
                case '"':
                    inQuotes = true;
                    break;
                case ',':
                    fields.Add(current.ToString());
                    current.Clear();
                    break;
                default:
                    current.Append(ch);
                    break;
            }
        }

        fields.Add(current.ToString());
        return fields;
    }
}
