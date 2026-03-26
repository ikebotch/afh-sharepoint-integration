namespace AFH.Common.SharePointUtils.Models;

public sealed class SharePointFieldDefinition
{
    public required string DisplayName { get; init; }
    public required string InternalName { get; init; }
    public required string Type { get; init; }
    public bool IsRequired { get; init; }
    public bool IsHidden { get; init; }
}
