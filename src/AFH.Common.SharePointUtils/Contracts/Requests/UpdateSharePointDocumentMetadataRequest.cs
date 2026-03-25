namespace AFH.Common.SharePointUtils.Contracts.Requests;

public sealed class UpdateSharePointDocumentMetadataRequest
{
    public required string SiteId { get; init; }
    public required string DriveId { get; init; }
    public required string ItemId { get; init; }
    public string? Name { get; init; }
}
