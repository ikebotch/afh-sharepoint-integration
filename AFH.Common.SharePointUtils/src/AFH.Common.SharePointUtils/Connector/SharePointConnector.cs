using AFH.Common.SharePointUtils.Connector.Interface;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Connector;

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
            return ConnectorResult<bool>.Fail(ex);
        }
    }

    public async Task<ConnectorResult<ListItemCollectionResponse>> GetListData(
      string siteId,
      string listId,
      string? filter = null,
      CancellationToken cancellationToken = default)
    {
        try
        {
            var items = await _graph.Sites[siteId]
                .Lists[listId]
                .Items
                .GetAsync(requestConfiguration =>
                {
                    if (!string.IsNullOrEmpty(filter))
                    {
                        requestConfiguration.QueryParameters.Filter = filter;
                    }
                }, cancellationToken);

            return ConnectorResult<ListItemCollectionResponse>.Ok(items!);
        }
        catch (Exception ex)
        {
            return ConnectorResult<ListItemCollectionResponse>.Fail(ex);
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
            return ConnectorResult<ListItem>.Fail(ex);
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
            return ConnectorResult<ListItem>.Fail(ex);
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
            return ConnectorResult<ListItem>.Fail(ex);
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
            return ConnectorResult<bool>.Fail(ex);
        }
    }
}
