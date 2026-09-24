namespace M365Manager.Core.SharedMailboxes;

public sealed class RenameSharedMailboxResult
{
    public required bool Succeeded { get; init; }

    public string? PreviousDisplayName { get; init; }
    public string? PreviousPrimarySmtpAddress { get; init; }

    public string? DisplayName { get; init; }
    public string? PrimarySmtpAddress { get; init; }

    /// <summary>One "MBX.Old.ED -&gt; MBX.New.ED" line per renamed access group, in ED/AU/RE order.</summary>
    public IReadOnlyList<string> RenamedGroups { get; init; } = Array.Empty<string>();

    public string? ErrorMessage { get; init; }

    /// <summary>
    /// Set when the rename itself went through but a follow-up step didn't - a group that couldn't
    /// be renamed, or the owner notification e-mail. <see cref="Succeeded"/> stays true; see
    /// <see cref="SharedMailboxCreationResult.WarningMessage"/>.
    /// </summary>
    public string? WarningMessage { get; init; }

    public Guid CorrelationId { get; init; }

    public static RenameSharedMailboxResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}

/// <summary>
/// What <see cref="ISharedMailboxService.PreviewRenameAsync"/> found for a mailbox: its current
/// name, address and access groups. The legacy dialog did the same thing in its second stage -
/// showed what the rename is about to touch before anything was changed.
/// </summary>
public sealed class SharedMailboxRenamePreview
{
    public required string DisplayName { get; init; }
    public required string PrimarySmtpAddress { get; init; }
    public required string Alias { get; init; }

    /// <summary>Domain part of <see cref="PrimarySmtpAddress"/> - the new address is built on it by default.</summary>
    public required string Domain { get; init; }

    /// <summary>Tier ("ED"/"AU"/"RE") -&gt; current group name, for whichever tiers this mailbox has.</summary>
    public required IReadOnlyDictionary<string, string> AccessGroups { get; init; }

    /// <summary>Display names of the current owner(s) - who the confirmation e-mail will go to. Empty when none could be resolved.</summary>
    public required string Owners { get; init; }
}
