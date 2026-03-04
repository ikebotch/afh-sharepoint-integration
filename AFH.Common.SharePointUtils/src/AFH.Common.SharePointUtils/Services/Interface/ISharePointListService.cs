using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Services.Interface;

public interface ISharePointListService
{
    Task<ListItem?> AddListItem(
        string? siteId,
        string? listId,
        IDictionary<string, object> fields,
        CancellationToken cancellationToken = default);

    Task<IReadOnlyList<ListItem>> GetListItems(
        string? siteId,
        string? listId,
        string? filter = null,
        string[]? selectFields = null,
        string[]? expandFields = null,
        CancellationToken cancellationToken = default);

    Task<ListItem?> GetListItemById(
        string? siteId,
        string? listId,
        string itemId,
        CancellationToken cancellationToken = default);

    Task<ListItem?> UpdateListItem(
        string? siteId,
        string? listId,
        string itemId,
        IDictionary<string, object> fields,
        CancellationToken cancellationToken = default);

    Task<bool> DeleteListItem(
        string? siteId,
        string? listId,
        string itemId,
        CancellationToken cancellationToken = default);

    Task<int> ClearListItems(string? siteId,
           string? listId,
        string? filter = null,
        CancellationToken cancellationToken = default);
}
