using AFH.Common.SharePointUtils.Connector.Interface;
using AFH.Common.SharePointUtils.Services.Interface;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Services.Implementation;

public class SharePointListService : ISharePointListService
{
    private readonly ISharePointConnector _connector;
    private readonly ILogger<SharePointListService> _logger;


    public SharePointListService(ISharePointConnector connector, ILogger<SharePointListService> logger,
        IConfiguration config)
    {
        _connector = connector;
        _logger = logger;
 
    }

    public async Task<ListItem?> AddListItem(
        string? siteId,
        string? listId,
        IDictionary<string, object> fields,
        CancellationToken cancellationToken = default)
    {
        var item = new ListItem
        {
            Fields = new FieldValueSet { AdditionalData = fields }
        };

        var result = await _connector.AddListItem(siteId!, listId!, item, cancellationToken);
        return result.Success ? result.Data : null;
    }

    public async Task<IReadOnlyList<ListItem>> GetListItems(
        string? siteId,
        string? listId,
        string? filter = null,
        string[]? selectFields = null,
        string[]? expandFields = null,
        CancellationToken cancellationToken = default)
    {
        var result = await _connector.GetListData(siteId!, listId!, filter, cancellationToken);

        return result.Success && result.Data?.Value != null
            ? result.Data.Value
            : Array.Empty<ListItem>();
    }

    public async Task<ListItem?> GetListItemById(
        string? siteId,
        string? listId,
        string itemId,
        CancellationToken cancellationToken = default)
    {
        var items = await GetListItems(siteId, listId, $"Id eq '{itemId}'", cancellationToken: cancellationToken);
        return items.FirstOrDefault();
    }

    public async Task<ListItem?> UpdateListItem(
        string? siteId,
        string? listId,
        string itemId,
        IDictionary<string, object> fields,
        CancellationToken cancellationToken = default)
    {
        var item = new ListItem
        {
            Fields = new FieldValueSet { AdditionalData = fields }
        };

        var result = await _connector.UpdateListItem(siteId!, listId!, itemId, item, cancellationToken);
        return result.Success ? result.Data : null;
    }

    public async Task<bool> DeleteListItem(
        string? siteId,
        string? listId,
        string itemId,
        CancellationToken cancellationToken = default)
    {
        var result = await _connector.DeleteListItem(siteId!, listId!, itemId, cancellationToken);
        return result.Success;
    }

    public async Task<int> ClearListItems(string? siteId,
        string? listId,
     string? filter = null,
     CancellationToken cancellationToken = default)
    {
        try
        {
            // Get items with optional filter
            var itemsResponse = await _connector.GetListData(
                siteId!,
                listId!,
                filter,
                cancellationToken);

            if (!itemsResponse.Success || itemsResponse.Data?.Value == null)
                return 0;

            int deletedCount = 0;

            foreach (var item in itemsResponse.Data.Value)
            {
                var deleteResult = await _connector.DeleteListItem(
                    siteId!,
                    listId!,
                    item.Id!,
                    cancellationToken);

                if (deleteResult.Success)
                    deletedCount++;
                else
                    _logger.LogWarning(deleteResult.Error, $"Failed to delete item {item.Id}");
            }

            return deletedCount;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error clearing list items");
            return 0;
        }
    }


}
