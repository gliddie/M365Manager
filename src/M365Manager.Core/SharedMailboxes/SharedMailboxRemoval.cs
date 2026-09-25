namespace M365Manager.Core.SharedMailboxes;

/// <summary>
/// What <see cref="ISharedMailboxService.PreviewRemovalAsync"/> found for a mailbox - the same
/// details the legacy ShrMbxRemove.ps1 dialog showed after "Get Shared Mailbox Details", so the
/// operator sees what goes before confirming.
/// </summary>
public sealed class SharedMailboxRemovalPreview
{
    public required string DisplayName { get; init; }
    public required string PrimarySmtpAddress { get; init; }

    /// <summary>Display names of the owner(s) - who the removal confirmation goes to. Empty when none resolved.</summary>
    public required string Owners { get; init; }

    /// <summary>ForwardingAddress / ForwardingSmtpAddress, whichever is set. Mail forwarded from here stops with the removal.</summary>
    public required string Forwarding { get; init; }

    public required IReadOnlyList<SharedMailboxAccessGroupInfo> AccessGroups { get; init; }

    /// <summary>
    /// People holding folder permissions directly rather than through an access group - the
    /// legacy dialog's "NonStandard Access" list. They lose access too, and nobody would know to
    /// tell them if the dialog didn't say so.
    /// </summary>
    public required IReadOnlyList<string> DirectAccess { get; init; }
}

public sealed class SharedMailboxAccessGroupInfo
{
    /// <summary>"ED", "AU" or "RE".</summary>
    public required string Tier { get; init; }
    public required string Name { get; init; }
    public required string Address { get; init; }

    /// <summary>Member display names, for showing.</summary>
    public required IReadOnlyList<string> MemberNames { get; init; }

    /// <summary>Member identities (primary SMTP address where there is one), for re-adding on recovery.</summary>
    public required IReadOnlyList<string> MemberIdentities { get; init; }
}

public sealed class RemoveSharedMailboxRequest
{
    /// <summary>ServiceNow (or equivalent) task number authorizing the removal. Required.</summary>
    public required string TaskNumber { get; init; }

    /// <summary>The mailbox: display name, alias, or any of its SMTP addresses.</summary>
    public required string MailboxIdentity { get; init; }

    /// <summary>
    /// Also delete the .ED/.AU/.RE groups. The legacy dialog had one checkbox per group, all
    /// ticked by default; the groups exist only for this mailbox, so one switch covers it.
    /// </summary>
    public bool DeleteAccessGroups { get; init; } = true;
}

public sealed class RemoveSharedMailboxResult
{
    public required bool Succeeded { get; init; }

    public string? DisplayName { get; init; }
    public string? PrimarySmtpAddress { get; init; }

    public IReadOnlyList<string> RemovedGroups { get; init; } = Array.Empty<string>();

    public string? ErrorMessage { get; init; }

    /// <summary>Set when the mailbox is gone but a follow-up step (a group, the e-mail) failed. <see cref="Succeeded"/> stays true.</summary>
    public string? WarningMessage { get; init; }

    public Guid CorrelationId { get; init; }

    public static RemoveSharedMailboxResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}

/// <summary>A soft-deleted shared mailbox - recoverable for 30 days after Remove-Mailbox.</summary>
public sealed class DeletedSharedMailbox
{
    public required Guid ExchangeGuid { get; init; }
    public required string DisplayName { get; init; }
    public required string PrimarySmtpAddress { get; init; }
    public DateTime? WhenSoftDeleted { get; init; }
}

/// <summary>
/// Written to the audit log (event code RSTDATA, target = ExchangeGuid) when the app removes a
/// shared mailbox, so a later recovery can put the owners and members back instead of asking the
/// operator to retype them from the ticket - the access groups are deleted for good, only the
/// mailbox itself comes back. The human-readable SNAP entry next to it is for people; this one is
/// for the recovery form.
/// </summary>
public sealed class SharedMailboxRestoreData
{
    public int Version { get; init; } = 1;
    public Guid ExchangeGuid { get; init; }
    public string DisplayName { get; init; } = "";
    public string PrimarySmtpAddress { get; init; } = "";
    public DateTime RemovedAtUtc { get; init; }
    public string TaskNumber { get; init; } = "";

    /// <summary>Owner UPNs, from the .ED group's ManagedBy.</summary>
    public List<string> Owners { get; init; } = new();

    public List<RestoreGroup> Groups { get; init; } = new();

    public sealed class RestoreGroup
    {
        public string Tier { get; init; } = "";
        public string Name { get; init; } = "";
        public List<string> Members { get; init; } = new();
    }
}

public sealed class RecoverSharedMailboxRequest
{
    public required string TaskNumber { get; init; }

    /// <summary>ExchangeGuid of the soft-deleted mailbox, from <see cref="DeletedSharedMailbox"/>.</summary>
    public required Guid ExchangeGuid { get; init; }

    /// <summary>Comma-separated owner identities for the recreated access groups. Required.</summary>
    public required string OwnerIdentities { get; init; }

    public string EditorMembers { get; init; } = "";
    public string AuthorMembers { get; init; } = "";
    public string ReaderMembers { get; init; } = "";
}

public sealed class RecoverSharedMailboxResult
{
    public required bool Succeeded { get; init; }

    public string? DisplayName { get; init; }
    public string? PrimarySmtpAddress { get; init; }

    /// <summary>".ED MBX.X.ED" per access group created (or reused) for the restored mailbox.</summary>
    public IReadOnlyList<string> AccessGroups { get; init; } = Array.Empty<string>();

    public string? ErrorMessage { get; init; }
    public string? WarningMessage { get; init; }
    public Guid CorrelationId { get; init; }

    public static RecoverSharedMailboxResult Failed(Guid correlationId, string error) => new()
    {
        Succeeded = false,
        ErrorMessage = error,
        CorrelationId = correlationId,
    };
}
