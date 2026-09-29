namespace M365Manager.Core.SharedMailboxes;

public enum AliasChangeMode
{
    Add,
    Remove,
}

/// <summary>A shared mailbox's SMTP addresses as they are now.</summary>
public sealed class SharedMailboxAddresses
{
    public required string DisplayName { get; init; }
    public required string PrimarySmtpAddress { get; init; }

    /// <summary>Secondary SMTP addresses, sorted. X500/SIP/SPO proxies are not listed - they are not e-mail aliases.</summary>
    public required IReadOnlyList<string> Aliases { get; init; }
}

public sealed class ChangeSharedMailboxAliasesRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>The mailbox: display name, alias, or any of its SMTP addresses.</summary>
    public required string MailboxIdentity { get; init; }

    public required AliasChangeMode Mode { get; init; }

    /// <summary>Full SMTP addresses, comma-separated.</summary>
    public required string Addresses { get; init; }

    /// <summary>
    /// Add mode, exactly one address: make it the primary address afterwards (the legacy script
    /// asked "Should this be made the new Primary SMTP Address?"). The old primary stays as an alias.
    /// </summary>
    public bool MakePrimary { get; init; }
}

public sealed class ChangeSharedMailboxAliasesResult
{
    public required bool Succeeded { get; init; }

    /// <summary>Addresses actually added or removed.</summary>
    public IReadOnlyList<string> Changed { get; init; } = Array.Empty<string>();

    /// <summary>Addresses with nothing to do - already there when adding, not there when removing.</summary>
    public IReadOnlyList<string> Skipped { get; init; } = Array.Empty<string>();

    /// <summary>"address: reason" for every address that failed.</summary>
    public IReadOnlyList<string> Failed { get; init; } = Array.Empty<string>();

    /// <summary>The primary address after the change.</summary>
    public string? PrimarySmtpAddress { get; init; }

    public string? ErrorMessage { get; init; }
    public Guid CorrelationId { get; init; }

    public static ChangeSharedMailboxAliasesResult Error(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}
