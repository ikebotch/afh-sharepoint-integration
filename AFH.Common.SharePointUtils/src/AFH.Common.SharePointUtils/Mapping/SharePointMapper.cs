using AFH.Common.SharePointUtils.Models;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Connector.Mapping;

public static class SharePointMapper
{
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
    {
        return new ListItem
        {
            Fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object>
                {
                    { "Title", dto.Title ?? string.Empty }
                }
            }
        };
    }
}
