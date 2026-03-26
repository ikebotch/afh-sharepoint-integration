namespace AFH.Common.SharePointUtils.Models;

public sealed class SharePointListItemModel
{
    public string? Id { get; init; }
    public DateTimeOffset? CreatedDateTime { get; init; }
    public DateTimeOffset? LastModifiedDateTime { get; init; }
    public IReadOnlyDictionary<string, object?> Fields { get; init; } = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
}
