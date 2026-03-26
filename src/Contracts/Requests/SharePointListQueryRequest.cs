namespace AFH.Common.SharePointUtils.Contracts.Requests;

public sealed class SharePointListQueryRequest
{
    public required string SiteId { get; init; }
    public required string ListId { get; init; }
    public string? Filter { get; init; }
    public string[]? SelectFields { get; init; }
    public string[]? ExpandFields { get; init; }
    public string[]? OrderBy { get; init; }
    public int? Top { get; init; }
    public string? SkipToken { get; init; }
}
