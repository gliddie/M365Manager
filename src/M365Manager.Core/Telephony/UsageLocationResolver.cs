using System.Globalization;
using System.Text.RegularExpressions;

namespace M365Manager.Core.Telephony;

/// <summary>
/// The usage location Microsoft 365 needs before a licence or a number can be assigned: an ISO 3166
/// two-letter country code. <see cref="Code"/> is null when it could not be worked out.
/// </summary>
/// <param name="Note">Why it was resolved this way, or what is wrong - shown on the review.</param>
/// <param name="NeedsCheck">The two sources disagree, or nothing was found.</param>
public sealed record UsageLocation(string? Code, string Note, bool NeedsCheck = false)
{
    public bool IsResolved => Code is not null;
}

/// <summary>
/// Turns a country as written in uchelper ("Germany", "USA", "United Kingdom") into the ISO code
/// Microsoft 365 accepts ("DE", "US", "GB"). Ported unchanged from Switchboard.
///
/// The Country column comes first, but it is not always maintained - so the country in the location
/// name ("Germany - Bremen") is the fallback, and a disagreement between the two is reported rather
/// than silently decided.
/// </summary>
public static class UsageLocationResolver
{
    /// <summary>Every country .NET knows, by ISO code, English and native name; built once.</summary>
    private static readonly Dictionary<string, string> ByName = BuildIndex();

    /// <summary>Spellings that do not match .NET's English region names.</summary>
    private static readonly Dictionary<string, string> Aliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["USA"] = "US", ["U.S.A."] = "US", ["United States of America"] = "US", ["America"] = "US",
        ["UK"] = "GB", ["U.K."] = "GB", ["Great Britain"] = "GB", ["England"] = "GB", ["Scotland"] = "GB", ["Wales"] = "GB", ["Northern Ireland"] = "GB",
        ["Korea"] = "KR", ["South Korea"] = "KR", ["Republic of Korea"] = "KR",
        ["Czech Republic"] = "CZ", ["Czechia"] = "CZ",
        ["Russia"] = "RU", ["Russian Federation"] = "RU",
        ["Vietnam"] = "VN", ["Viet Nam"] = "VN",
        ["Turkey"] = "TR", ["Türkiye"] = "TR", ["Turkiye"] = "TR",
        ["Holland"] = "NL", ["The Netherlands"] = "NL",
        ["UAE"] = "AE", ["United Arab Emirates"] = "AE",
        ["Hong Kong"] = "HK", ["Hong Kong SAR"] = "HK",
        ["Taiwan"] = "TW", ["Macau"] = "MO", ["Macao"] = "MO",
        ["Deutschland"] = "DE", ["PRC"] = "CN", ["China"] = "CN",
    };

    public static UsageLocation Resolve(string? country, string? locationName)
    {
        // The column sometimes carries a region too ("Canada - ON") - cut it like a location name.
        var fromColumn = Lookup(country) ?? Lookup(CountryPartOf(country));
        var namePart = CountryPartOf(locationName);
        var fromName = Lookup(namePart);

        if (fromColumn is not null && fromName is not null && fromColumn != fromName)
            return new UsageLocation(fromColumn,
                $"Country column says {country} ({fromColumn}), the location name says {namePart} ({fromName}) - check which is right.",
                NeedsCheck: true);

        if (fromColumn is not null)
            return new UsageLocation(fromColumn, $"From the Country column ({country}).");

        if (fromName is not null)
            return new UsageLocation(fromName,
                string.IsNullOrWhiteSpace(country)
                    ? $"From the location name ({namePart}) - the Country column is empty."
                    : $"From the location name ({namePart}) - the Country column '{country}' is not a known country.");

        return new UsageLocation(null,
            $"No country found in the Country column ('{country}') or the location name ('{locationName}').",
            NeedsCheck: true);
    }

    /// <summary>
    /// The country part of a location name as uchelper writes them: "Germany - Bremen" → "Germany",
    /// "Netherlands -- Arnhem" → "Netherlands", "Singapore (Lab)" → "Singapore".
    /// </summary>
    public static string CountryPartOf(string? locationName)
    {
        var name = (locationName ?? "").Trim();
        name = Regex.Split(name, @"\s+-+\s*")[0];
        name = Regex.Replace(name, @"\s*\(.*\)\s*$", "");
        return name.Trim();
    }

    /// <summary>ISO code for a country name or code, or null.</summary>
    public static string? Lookup(string? value)
    {
        var key = (value ?? "").Trim();
        if (key.Length == 0)
            return null;
        if (Aliases.TryGetValue(key, out var alias))
            return alias;
        return ByName.TryGetValue(key, out var code) ? code : null;
    }

    private static Dictionary<string, string> BuildIndex()
    {
        var index = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var culture in CultureInfo.GetCultures(CultureTypes.SpecificCultures))
        {
            RegionInfo region;
            try { region = new RegionInfo(culture.Name); }
            catch { continue; }

            var iso = region.TwoLetterISORegionName;
            if (iso.Length != 2 || !char.IsLetter(iso[0]))
                continue;

            index.TryAdd(iso, iso);
            index.TryAdd(region.ThreeLetterISORegionName, iso);
            index.TryAdd(region.EnglishName, iso);
            index.TryAdd(region.NativeName, iso);
        }
        return index;
    }
}
