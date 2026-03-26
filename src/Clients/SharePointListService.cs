using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Contracts.Requests;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Clients;

public class SharePointListService : ISharePointListService
{
    private readonly ISharePointListClient _listClient;
    private readonly ILogger<SharePointListService> _logger;

    public SharePointListService(
        ISharePointListClient listClient,
        ILogger<SharePointListService> logger)
    {
        _listClient = listClient;
        _logger = logger;
    }

    public async Task<ListItem?> AddListItem(
        string? siteId,
        string? listId,
        IDictionary<string, object> fields,
        CancellationToken cancellationToken = default)
    {
        var result = await _listClient.CreateListItemAsync(new SharePointListItemUpsertRequest
        {
            SiteId = siteId!,
            ListId = listId!,
            Fields = fields.ToDictionary(kvp => kvp.Key, kvp => (object?)kvp.Value)
        }, cancellationToken);

        return result is null ? null : ToGraphItem(result);
    }

    public async Task<IReadOnlyList<ListItem>> GetListItems(
        string? siteId,
        string? listId,
        string? filter = null,
        string[]? selectFields = null,
        string[]? expandFields = null,
        CancellationToken cancellationToken = default)
    {
        var result = await _listClient.QueryListItemsAsync(new SharePointListQueryRequest
        {
            SiteId = siteId!,
            ListId = listId!,
            Filter = filter,
            SelectFields = selectFields,
            ExpandFields = expandFields
        }, cancellationToken);

        return result.Items.Select(ToGraphItem).ToList();
    }

    public async Task<ListItem?> GetListItemById(
        string? siteId,
        string? listId,
        string itemId,
        CancellationToken cancellationToken = default)
    {
        var result = await _listClient.GetListItemAsync(siteId!, listId!, itemId, cancellationToken);
        return result is null ? null : ToGraphItem(result);
    }

    public async Task<ListItem?> UpdateListItem(
        string? siteId,
        string? listId,
        string itemId,
        IDictionary<string, object> fields,
        CancellationToken cancellationToken = default)
    {
        var result = await _listClient.UpdateListItemAsync(new SharePointListItemUpsertRequest
        {
            SiteId = siteId!,
            ListId = listId!,
            ItemId = itemId,
            Fields = fields.ToDictionary(kvp => kvp.Key, kvp => (object?)kvp.Value)
        }, cancellationToken);

        return result is null ? null : ToGraphItem(result);
    }

    public async Task<bool> DeleteListItem(
        string? siteId,
        string? listId,
        string itemId,
        CancellationToken cancellationToken = default)
    {
        await _listClient.DeleteListItemAsync(siteId!, listId!, itemId, cancellationToken);
        return true;
    }

    public async Task<int> ClearListItems(
        string? siteId,
        string? listId,
        string? filter = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            var itemsResponse = await _listClient.QueryListItemsAsync(new SharePointListQueryRequest
            {
                SiteId = siteId!,
                ListId = listId!,
                Filter = filter
            }, cancellationToken);

            var deletedCount = 0;

            foreach (var item in itemsResponse.Items)
            {
                try
                {
                    await _listClient.DeleteListItemAsync(siteId!, listId!, item.Id!, cancellationToken);
                    deletedCount++;
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed to delete item {ItemId}", item.Id);
                }
            }

            return deletedCount;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error clearing list items");
            return 0;
        }
    }

    private static ListItem ToGraphItem(Models.SharePointListItemModel item)
        => new()
        {
            Id = item.Id,
            CreatedDateTime = item.CreatedDateTime,
            LastModifiedDateTime = item.LastModifiedDateTime,
            Fields = new FieldValueSet
            {
                AdditionalData = item.Fields.ToDictionary(kvp => kvp.Key, kvp => kvp.Value)
            }
        };
}
