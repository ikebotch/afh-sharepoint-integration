using AFH.Common.SharePointUtils.Contracts.Requests;
using AFH.Common.SharePointUtils.Contracts.Responses;
using AFH.Common.SharePointUtils.Models;

namespace AFH.Common.SharePointUtils.Abstractions;

public interface ISharePointListClient
{
    Task<bool> ListExistsAsync(string siteId, string listId, CancellationToken cancellationToken = default);
    Task<SharePointListItemsResponse> QueryListItemsAsync(SharePointListQueryRequest request, CancellationToken cancellationToken = default);
    Task<SharePointListItemModel?> GetListItemAsync(string siteId, string listId, string itemId, CancellationToken cancellationToken = default);
    Task<SharePointListItemModel?> CreateListItemAsync(SharePointListItemUpsertRequest request, CancellationToken cancellationToken = default);
    Task<SharePointListItemModel?> UpdateListItemAsync(SharePointListItemUpsertRequest request, CancellationToken cancellationToken = default);
    Task DeleteListItemAsync(string siteId, string listId, string itemId, CancellationToken cancellationToken = default);
}
