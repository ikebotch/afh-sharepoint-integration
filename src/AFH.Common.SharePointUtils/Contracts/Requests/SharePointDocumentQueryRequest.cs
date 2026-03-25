namespace AFH.Common.SharePointUtils.Contracts.Requests;

public sealed class SharePointDocumentQueryRequest
{
    public required string SiteId { get; init; }
    public required string DriveId { get; init; }
    public string? FolderItemId { get; init; }
}
