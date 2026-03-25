using AFH.Common.SharePointUtils.Models;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Mapping;

public static class SharePointMapper
{
    public static SharePointListItemModel ToModel(ListItem item)
    {
        var fields = item.Fields?.AdditionalData is null
            ? new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, object?>(item.Fields.AdditionalData.ToDictionary(kvp => kvp.Key, kvp => (object?)kvp.Value), StringComparer.OrdinalIgnoreCase);

        return new SharePointListItemModel
        {
            Id = item.Id,
            CreatedDateTime = item.CreatedDateTime,
            LastModifiedDateTime = item.LastModifiedDateTime,
            Fields = fields
        };
    }

    public static SharePointListItemDto ToDto(ListItem item)
    {
        var fields = item.Fields?.AdditionalData;

        return new SharePointListItemDto
        {
            Id = item.Id,
            Title = fields != null && fields.TryGetValue("Title", out var t) ? t?.ToString() : null,
            Created = item.CreatedDateTime,
            Modified = item.LastModifiedDateTime
        };
    }

    public static ListItem ToListItem(SharePointListItemDto dto)
        => ToListItem(new Dictionary<string, object?>
        {
            ["Title"] = dto.Title ?? string.Empty
        });

    public static ListItem ToListItem(IReadOnlyDictionary<string, object?> fields)
    {
        return new ListItem
        {
            Fields = new FieldValueSet
            {
                AdditionalData = fields.ToDictionary(kvp => kvp.Key, kvp => kvp.Value)
            }
        };
    }
}
