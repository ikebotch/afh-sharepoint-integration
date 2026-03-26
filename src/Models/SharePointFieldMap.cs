namespace AFH.Common.SharePointUtils.Models;

public sealed class SharePointFieldMap
{
    public required string PropertyName { get; init; }
    public required string FieldName { get; init; }
    public SharePointFieldNameKind NameKind { get; init; } = SharePointFieldNameKind.Auto;
}
