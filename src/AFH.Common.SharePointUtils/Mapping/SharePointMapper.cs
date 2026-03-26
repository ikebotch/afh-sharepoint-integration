using AFH.Common.SharePointUtils.Extensions;
using AFH.Common.SharePointUtils.Models;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Mapping;

public static class SharePointMapper
{
    public static SharePointListItemModel ToModel(ListItem item)
    {
        return new SharePointListItemModel
        {
            Id = item.Id,
            CreatedDateTime = item.CreatedDateTime,
            LastModifiedDateTime = item.LastModifiedDateTime,
            Fields = new Dictionary<string, object?>(item.GetFieldValues(), StringComparer.OrdinalIgnoreCase)
        };
    }

    public static SharePointListItemDto ToDto(ListItem item)
    {
        var fields = item.GetFieldValues();

        return new SharePointListItemDto
        {
            Id = item.Id,
            Title = fields.GetString("Title"),
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
