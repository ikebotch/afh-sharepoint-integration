using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Contracts.Requests;
using AFH.Common.SharePointUtils.Contracts.Responses;
using AFH.Common.SharePointUtils.Models;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Clients;

public sealed class GraphSharePointDocumentClient : ISharePointDocumentClient
{
    private readonly GraphServiceClient _graph;

    public GraphSharePointDocumentClient(GraphServiceClient graph)
    {
        _graph = graph;
    }

    public async Task<IReadOnlyList<SharePointDocumentMetadata>> ListDocumentsAsync(SharePointDocumentQueryRequest request, CancellationToken cancellationToken = default)
    {
        DriveItemCollectionResponse? response;

        if (string.IsNullOrWhiteSpace(request.FolderItemId))
        {
            var root = await _graph.Drives[request.DriveId].Root.GetAsync(cancellationToken: cancellationToken);
            if (root?.Id is null)
                return [];

            response = await _graph.Drives[request.DriveId].Items[root.Id].Children.GetAsync(cancellationToken: cancellationToken);
        }
        else
        {
            response = await _graph.Drives[request.DriveId].Items[request.FolderItemId].Children.GetAsync(cancellationToken: cancellationToken);
        }

        return response?.Value?.Select(MapDocument).ToList() ?? [];
    }

    public async Task<SharePointDocumentMetadata?> GetDocumentMetadataAsync(string siteId, string driveId, string itemId, CancellationToken cancellationToken = default)
    {
        var item = await _graph.Drives[driveId].Items[itemId].GetAsync(cancellationToken: cancellationToken);
        return item is null ? null : MapDocument(item);
    }

    public async Task<SharePointDocumentContentResponse?> DownloadDocumentAsync(string siteId, string driveId, string itemId, CancellationToken cancellationToken = default)
    {
        var metadata = await GetDocumentMetadataAsync(siteId, driveId, itemId, cancellationToken);
        if (metadata is null)
            return null;

        var content = await _graph.Drives[driveId].Items[itemId].Content.GetAsync(cancellationToken: cancellationToken);
        if (content is null)
            return null;

        return new SharePointDocumentContentResponse
        {
            Metadata = metadata,
            Content = content
        };
    }

    public async Task<SharePointDocumentMetadata?> UploadDocumentAsync(SharePointDocumentUploadRequest request, CancellationToken cancellationToken = default)
    {
        var item = await _graph.Drives[request.DriveId]
            .Root
            .ItemWithPath(request.FilePath)
            .Content
            .PutAsync(request.Content, cancellationToken: cancellationToken);

        return item is null ? null : MapDocument(item);
    }

    public async Task<SharePointDocumentMetadata?> UpdateDocumentMetadataAsync(UpdateSharePointDocumentMetadataRequest request, CancellationToken cancellationToken = default)
    {
        var updated = await _graph.Drives[request.DriveId]
            .Items[request.ItemId]
            .PatchAsync(new DriveItem
            {
                Name = request.Name
            }, cancellationToken: cancellationToken);

        return updated is null ? null : MapDocument(updated);
    }

    public Task DeleteDocumentAsync(string siteId, string driveId, string itemId, CancellationToken cancellationToken = default)
        => _graph.Drives[driveId].Items[itemId].DeleteAsync(cancellationToken: cancellationToken);

    private static SharePointDocumentMetadata MapDocument(DriveItem item)
        => new()
        {
            ItemId = item.Id,
            Name = item.Name,
            WebUrl = item.WebUrl,
            Size = item.Size,
            LastModifiedDateTime = item.LastModifiedDateTime,
            ParentPath = item.ParentReference?.Path,
            DownloadUrl = item.AdditionalData != null && item.AdditionalData.TryGetValue("@microsoft.graph.downloadUrl", out var value)
                ? value?.ToString()
                : null
        };
}
