using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Contracts.Requests;
using AFH.Common.SharePointUtils.Contracts.Responses;
using AFH.Common.SharePointUtils.Exceptions;
using AFH.Common.SharePointUtils.Mapping;
using AFH.Common.SharePointUtils.Models;

namespace AFH.Common.SharePointUtils.Clients;

public sealed class GraphSharePointListClient : ISharePointListClient
{
    private readonly ISharePointConnector _connector;

    public GraphSharePointListClient(ISharePointConnector connector)
    {
        _connector = connector;
    }

    public async Task<bool> ListExistsAsync(string siteId, string listId, CancellationToken cancellationToken = default)
    {
        var result = await _connector.ListExists(siteId, listId, cancellationToken);
        if (result.Success)
            return result.Data;

        throw result.Error ?? new SharePointSdkException("Failed to verify SharePoint list existence.");
    }

    public async Task<SharePointListItemsResponse> QueryListItemsAsync(SharePointListQueryRequest request, CancellationToken cancellationToken = default)
    {
        var result = await _connector.GetListData(request, cancellationToken);
        if (!result.Success || result.Data is null)
            throw result.Error ?? new SharePointSdkException("Failed to query SharePoint list items.");

        return new SharePointListItemsResponse
        {
            Items = result.Data.Value?.Select(SharePointMapper.ToModel).ToList() ?? [],
            NextPageLink = result.Data.OdataNextLink
        };
    }

    public async Task<SharePointListItemModel?> GetListItemAsync(string siteId, string listId, string itemId, CancellationToken cancellationToken = default)
    {
        var result = await _connector.GetListItem(siteId, listId, itemId, cancellationToken);
        if (!result.Success)
            throw result.Error ?? new SharePointSdkException("Failed to get SharePoint list item.");

        return result.Data is null ? null : SharePointMapper.ToModel(result.Data);
    }

    public async Task<SharePointListItemModel?> CreateListItemAsync(SharePointListItemUpsertRequest request, CancellationToken cancellationToken = default)
    {
        var result = await _connector.AddListItem(request.SiteId, request.ListId, SharePointMapper.ToListItem(request.Fields), cancellationToken);
        if (!result.Success)
            throw result.Error ?? new SharePointSdkException("Failed to create SharePoint list item.");

        return result.Data is null ? null : SharePointMapper.ToModel(result.Data);
    }

    public async Task<SharePointListItemModel?> UpdateListItemAsync(SharePointListItemUpsertRequest request, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(request.ItemId))
            throw new ArgumentException("ItemId is required for update operations.", nameof(request));

        var result = await _connector.UpdateListItem(request.SiteId, request.ListId, request.ItemId, SharePointMapper.ToListItem(request.Fields), cancellationToken);
        if (!result.Success)
            throw result.Error ?? new SharePointSdkException("Failed to update SharePoint list item.");

        return result.Data is null ? null : SharePointMapper.ToModel(result.Data);
    }

    public async Task DeleteListItemAsync(string siteId, string listId, string itemId, CancellationToken cancellationToken = default)
    {
        var result = await _connector.DeleteListItem(siteId, listId, itemId, cancellationToken);
        if (!result.Success)
            throw result.Error ?? new SharePointSdkException("Failed to delete SharePoint list item.");
    }
}
