using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Contracts.Requests;
using AFH.Common.SharePointUtils.Contracts.Responses;
using AFH.Common.SharePointUtils.Exceptions;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Clients;

public class SharePointConnector : ISharePointConnector
{
    private readonly GraphServiceClient _graph;

    public SharePointConnector(GraphServiceClient graph)
    {
        _graph = graph;
    }

    public async Task<ConnectorResult<bool>> ListExists(string siteId, string listId, CancellationToken cancellationToken = default)
    {
        try
        {
            var list = await _graph.Sites[siteId].Lists[listId].GetAsync(cancellationToken: cancellationToken);
            return ConnectorResult<bool>.Ok(list != null);
        }
        catch (Exception ex)
        {
            return ConnectorResult<bool>.Fail(new SharePointSdkException("Failed to verify SharePoint list existence.", ex));
        }
    }

    public async Task<ConnectorResult<ListItemCollectionResponse>> GetListData(
        SharePointListQueryRequest request,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var items = await _graph.Sites[request.SiteId]
                .Lists[request.ListId]
                .Items
                .GetAsync(requestConfiguration =>
                {
                    if (!string.IsNullOrWhiteSpace(request.Filter))
                        requestConfiguration.QueryParameters.Filter = request.Filter;

                    if (request.ExpandFields is { Length: > 0 })
                        requestConfiguration.QueryParameters.Expand = request.ExpandFields;

                    if (request.OrderBy is { Length: > 0 })
                        requestConfiguration.QueryParameters.Orderby = request.OrderBy;

                    if (request.SelectFields is { Length: > 0 })
                        requestConfiguration.QueryParameters.Select = request.SelectFields;

                    if (request.Top is > 0)
                        requestConfiguration.QueryParameters.Top = request.Top;

                }, cancellationToken);

            return ConnectorResult<ListItemCollectionResponse>.Ok(items!);
        }
        catch (Exception ex)
        {
            return ConnectorResult<ListItemCollectionResponse>.Fail(
                new SharePointSdkException("Failed to query SharePoint list items.", ex));
        }
    }

    public async Task<ConnectorResult<ListItem>> GetListItem(string siteId, string listId, string itemId, CancellationToken cancellationToken = default)
    {
        try
        {
            var item = await _graph.Sites[siteId].Lists[listId].Items[itemId].GetAsync(cancellationToken: cancellationToken);
            return ConnectorResult<ListItem>.Ok(item!);
        }
        catch (Exception ex)
        {
            return ConnectorResult<ListItem>.Fail(new SharePointSdkException("Failed to get SharePoint list item.", ex));
        }
    }

    public async Task<ConnectorResult<ListItem>> AddListItem(string siteId, string listId, ListItem item, CancellationToken cancellationToken = default)
    {
        try
        {
            var added = await _graph.Sites[siteId].Lists[listId].Items.PostAsync(item, cancellationToken: cancellationToken);
            return ConnectorResult<ListItem>.Ok(added!);
        }
        catch (Exception ex)
        {
            return ConnectorResult<ListItem>.Fail(new SharePointSdkException("Failed to create SharePoint list item.", ex));
        }
    }

    public async Task<ConnectorResult<ListItem>> UpdateListItem(string siteId, string listId, string itemId, ListItem item, CancellationToken cancellationToken = default)
    {
        try
        {
            var updated = await _graph.Sites[siteId].Lists[listId].Items[itemId].PatchAsync(item, cancellationToken: cancellationToken);
            return ConnectorResult<ListItem>.Ok(updated!);
        }
        catch (Exception ex)
        {
            return ConnectorResult<ListItem>.Fail(new SharePointSdkException("Failed to update SharePoint list item.", ex));
        }
    }

    public async Task<ConnectorResult<bool>> DeleteListItem(string siteId, string listId, string itemId, CancellationToken cancellationToken = default)
    {
        try
        {
            await _graph.Sites[siteId].Lists[listId].Items[itemId].DeleteAsync(cancellationToken: cancellationToken);
            return ConnectorResult<bool>.Ok(true);
        }
        catch (Exception ex)
        {
            return ConnectorResult<bool>.Fail(new SharePointSdkException("Failed to delete SharePoint list item.", ex));
        }
    }
}
