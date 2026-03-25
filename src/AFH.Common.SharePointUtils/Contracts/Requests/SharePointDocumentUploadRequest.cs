namespace AFH.Common.SharePointUtils.Contracts.Requests;

public sealed class SharePointDocumentUploadRequest
{
    public required string SiteId { get; init; }
    public required string DriveId { get; init; }
    public required string FilePath { get; init; }
    public required Stream Content { get; init; }
    public string? ContentType { get; init; }
}
