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

    /// <summary>Owners (ManagedBy), display-friendly.</summary>
    public List<string> ManagedBy { get; set; } = new();
}

/// <summary>A member of a group (from Get-DistributionGroupMember).</summary>
public sealed class GroupMemberInfo
{
    public string DisplayName { get; set; } = "";
    public string PrimarySmtpAddress { get; set; } = "";
    public string RecipientType { get; set; } = "";
}
