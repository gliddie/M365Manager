using System.Text.RegularExpressions;
using M365Manager.Core.Naming;
using M365Manager.Data.Naming;

namespace M365Manager.Core.Teams;

public sealed partial class TeamNamingService : ITeamNamingService
{
    private readonly ITeamNamingRepository _repository;

    public TeamNamingService(ITeamNamingRepository repository)
    {
        _repository = repository;
    }

    public async Task<string> BuildDisplayNameAsync(string location, string name, CancellationToken ct = default)
    {
        var acronyms = await _repository.GetAllAsync(ct);

        // Case-insensitive, like the other naming services - a plain string.Replace matches
        // nothing once casing differs and leaves the acronym table inert. See AcronymSubstitution.
        var proposed = AcronymSubstitution.Apply((name ?? "").Trim(), acronyms);

        return $"GRP.{(location ?? "").Trim()}.{proposed}";
    }

    public string BuildAlias(string displayName)
    {
        // The legacy character class, then the shared sanitizer, so "&" or an umlaut can't produce
        // a mailNickname Graph rejects.
        var alias = ExchangeNaming.ToAliasSafe(InvalidAliasCharsRegex().Replace(displayName ?? "", ""));
        return alias.Length > 64 ? alias[..64] : alias;
    }

    // Mirrors the legacy: -Replace '[ /,$#_-]',''  then .Replace("\","").Replace(".","").Replace(" ","")
    [GeneratedRegex(@"[ /,$#_\-\\.]")]
    private static partial Regex InvalidAliasCharsRegex();
}
