namespace M365Manager.Core.Exchange;

/// <summary>Summary of a distribution / security group (from Get-DistributionGroup).</summary>
public sealed class DistributionGroupInfo
{
    public string Name { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string PrimarySmtpAddress { get; set; } = "";
    public string Alias { get; set; } = "";
    public string GroupType { get; set; } = "";
    public string RecipientTypeDetails { get; set; } = "";
    public string Notes { get; set; } = "";
    public bool IsDeletedInM365 { get; set; }
    public DateTime? LastSeenAtUtc { get; set; }

    /// <summary>
    /// Entra (Graph) object id. Cached (SQL) lookups leave this blank - <see cref="Name"/> already
    /// IS the Graph id there. Live Exchange lookups populate it from ExternalDirectoryObjectId,
    /// since Name is just the Exchange recipient name in that case.
    /// </summary>
    public string ExternalDirectoryObjectId { get; set; } = "";

    /// <summary>From Exchange (live lookups) or Graph's mailEnabled/securityEnabled (cached lookups).</summary>
    public bool MailEnabled { get; set; }
    public bool SecurityEnabled { get; set; }
    public bool IsMembershipDynamic { get; set; }

    public string StatusText => IsDeletedInM365 ? "Deleted in M365" : "";

    /// <summary>Owners (ManagedBy), display-friendly.</summary>
    public List<string> ManagedBy { get; set; } = new();

    /// <summary>
    /// Short, human-friendly group type label. Live Exchange lookups populate
    /// <see cref="RecipientTypeDetails"/>, which is preferred when present. Cached (Graph-sourced)
    /// data has no RecipientTypeDetails - Graph's group resource doesn't have that concept - so it
    /// falls back to GroupTypes/MailEnabled/SecurityEnabled instead.
    /// </summary>
    public string GroupTypeLabel
    {
        get
        {
            var details = RecipientTypeDetails ?? "";
            if (details.Contains("GroupMailbox", StringComparison.OrdinalIgnoreCase))
                return "Microsoft 365";
            if (details.Contains("MailUniversalSecurityGroup", StringComparison.OrdinalIgnoreCase))
                return "Security";
            if (details.Contains("DynamicDistributionGroup", StringComparison.OrdinalIgnoreCase))
                return "Dynamic";
            if (details.Contains("MailUniversalDistributionGroup", StringComparison.OrdinalIgnoreCase)
                || details.Contains("MailNonUniversalGroup", StringComparison.OrdinalIgnoreCase))
                return "Distribution";
            if (!string.IsNullOrWhiteSpace(details))
                return details;

            if ((GroupType ?? "").Contains("Unified", StringComparison.OrdinalIgnoreCase))
                return "Microsoft 365";
            if (IsMembershipDynamic)
                return "Dynamic Security";
            if (MailEnabled && SecurityEnabled)
                return "Mail-enabled Security";
            if (MailEnabled)
                return "Distribution";
            if (SecurityEnabled)
                return "Security";

            return "Group";
        }
    }

    /// <summary>Category key used to pick a badge color for <see cref="GroupTypeLabel"/>.</summary>
    public string GroupTypeKey => GroupTypeLabel switch
    {
        "Microsoft 365" => "M365Group",
        "Security" or "Mail-enabled Security" => "SecurityGroup",
        "Dynamic" or "Dynamic Security" => "DynamicGroup",
        "Distribution" => "DistributionGroup",
        _ => "OtherGroup",
    };

    /// <summary>
    /// Whether this is a Microsoft 365 (Unified) group, which uses Get-UnifiedGroupLinks /
    /// Add-UnifiedGroupLinks instead of the DistributionGroupMember cmdlets. Derived from
    /// <see cref="GroupTypeLabel"/> so it works for both live (RecipientTypeDetails) and cached
    /// (GroupTypes/MailEnabled/SecurityEnabled) data.
    /// </summary>
    public bool IsUnifiedGroup => GroupTypeLabel == "Microsoft 365";
}

/// <summary>A member of a group (from Get-DistributionGroupMember).</summary>
public sealed class GroupMemberInfo
{
    public string DisplayName { get; set; } = "";
    public string PrimarySmtpAddress { get; set; } = "";
    public string UserPrincipalName { get; set; } = "";
    public string RecipientType { get; set; } = "";
    public bool IsNested { get; set; }
    public string Path { get; set; } = "";
}

/// <summary>Outcome of adding/removing one member as part of a (possibly bulk) operation.</summary>
public sealed record MemberOperationResult(string Identity, bool Succeeded, string? Error);
