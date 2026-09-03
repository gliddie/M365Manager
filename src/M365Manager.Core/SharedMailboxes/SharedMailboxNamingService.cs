using System.Globalization;
using M365Manager.Core.Naming;
using M365Manager.Data.Naming;

namespace M365Manager.Core.SharedMailboxes;

/// <summary>
/// See <see cref="ISharedMailboxNamingService"/>. Reuses <see cref="ITeamNamingRepository"/> as-is
/// (dbo.TeamNameAcronyms) rather than a separate table - the legacy scripts sourced both Teams and
/// Shared Mailbox naming from the same KnownAcronyms.csv.
/// </summary>
public sealed class SharedMailboxNamingService : ISharedMailboxNamingService
{
    private readonly ITeamNamingRepository _repository;

    public SharedMailboxNamingService(ITeamNamingRepository repository)
    {
        _repository = repository;
    }

    public async Task<(string DisplayName, string LocalPart)> BuildNameAsync(string location, string name, CancellationToken ct = default)
    {
        var loc = (location ?? "").Trim();

        // The name is used exactly as typed. The legacy scripts ran it through ToTitleCase first,
        // which mangles acronyms ("ULS M&A IT Support" -> "Uls M&A It Support") - and then needed
        // 103 of their 117 KnownAcronyms entries to be identity mappings (CRS -> CRS, HVAC -> HVAC)
        // purely to undo that damage. Keeping the requester's casing removes the whole problem;
        // the acronym table is left for genuine translations (R&I -> RI).
        var acronyms = await _repository.GetAllAsync(ct);
        var proposed = AcronymSubstitution.Apply((name ?? "").Trim(), acronyms);

        // The location IS normalized: it's a site code, and "Global*" is title-cased rather than
        // upper-cased - both are real conventions rather than an artifact, so they stay.
        var locFormatted = loc.StartsWith("Global", StringComparison.OrdinalIgnoreCase)
            ? TitleCase(loc)
            : loc.ToUpperInvariant();

        // The display name keeps whatever was requested ("ULS M&A IT Support"); only the address
        // is reduced to characters Exchange accepts.
        var displayName = $"{locFormatted} {proposed}";
        var localPart = ExchangeNaming.ToAliasSafe($"{locFormatted}.{proposed.Replace(" ", "")}");

        // Exchange caps an alias at 64 characters.
        if (localPart.Length > AliasLimit)
            localPart = localPart[..AliasLimit].TrimEnd('.');

        return (displayName, localPart);
    }

    public string BuildAddress(string localPart, string domain) => $"{localPart}@{domain}";

    public string BuildAccessGroupName(string mailboxDisplayName, string tier)
    {
        var name = $"MBX.{(mailboxDisplayName ?? "").Trim()}.{tier}";

        // Exchange caps a group Name at 64 characters too. Trim the middle rather than let
        // New-DistributionGroup fail outright on an over-long mailbox name.
        if (name.Length <= NameLimit)
            return name;

        var suffix = $".{tier}";
        return name[..(NameLimit - suffix.Length)].TrimEnd() + suffix;
    }

    public string BuildAccessGroupAlias(string mailboxDisplayName, string tier)
    {
        // The spaces go, the dots around "MBX." and the tier stay, and anything Exchange won't
        // take in an alias (&, umlauts, apostrophes, ...) is stripped the same way as for the
        // mailbox address itself.
        var alias = ExchangeNaming.ToAliasSafe(BuildAccessGroupName(mailboxDisplayName, tier).Replace(" ", ""));
        return alias.Length > AliasLimit ? alias[..AliasLimit].TrimEnd('.') : alias;
    }

    public string BuildAccessGroupAddress(string mailboxDisplayName, string tier, string domain)
        => $"{BuildAccessGroupAlias(mailboxDisplayName, tier)}@{domain}";

    private const int AliasLimit = 64;
    private const int NameLimit = 64;

    private static string TitleCase(string value) => CultureInfo.InvariantCulture.TextInfo.ToTitleCase(value.ToLowerInvariant());
}
