namespace AFH.Common.SharePointUtils.Models;

public sealed class SharePointDocumentMetadata
{
    public string? ItemId { get; init; }
    public string? Name { get; init; }
    public string? WebUrl { get; init; }
    public long? Size { get; init; }
    public DateTimeOffset? LastModifiedDateTime { get; init; }
    public string? DownloadUrl { get; init; }
    public string? ParentPath { get; init; }
}
