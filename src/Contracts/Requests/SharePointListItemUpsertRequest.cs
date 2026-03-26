namespace AFH.Common.SharePointUtils.Contracts.Requests;

public sealed class SharePointListItemUpsertRequest
{
    public required string SiteId { get; init; }
    public required string ListId { get; init; }
    public string? ItemId { get; init; }
    public IReadOnlyDictionary<string, object?> Fields { get; init; } = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
}
