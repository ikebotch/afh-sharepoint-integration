using AFH.Common.SharePointUtils.Contracts.Requests;
using AFH.Common.SharePointUtils.Contracts.Responses;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Abstractions;

public interface ISharePointConnector
{
    Task<ConnectorResult<bool>> ListExists(string siteId, string listId, CancellationToken cancellationToken = default);

    Task<ConnectorResult<ListItemCollectionResponse>> GetListData(
        SharePointListQueryRequest request,
        CancellationToken cancellationToken = default);

    Task<ConnectorResult<ListItem>> GetListItem(
        string siteId,
        string listId,
        string itemId,
        CancellationToken cancellationToken = default);

    Task<ConnectorResult<ListItem>> AddListItem(
        string siteId,
        string listId,
        ListItem item,
        CancellationToken cancellationToken = default);

    Task<ConnectorResult<ListItem>> UpdateListItem(
        string siteId,
        string listId,
        string itemId,
        ListItem item,
        CancellationToken cancellationToken = default);

    Task<ConnectorResult<bool>> DeleteListItem(
        string siteId,
        string listId,
        string itemId,
        CancellationToken cancellationToken = default);
}
