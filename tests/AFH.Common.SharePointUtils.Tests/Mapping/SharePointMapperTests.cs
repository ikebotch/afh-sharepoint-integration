using AFH.Common.SharePointUtils.Mapping;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Tests.Mapping;

public class SharePointMapperTests
{
    [Fact]
    public void ToModel_Maps_Graph_ListItem_Into_Generic_Field_Bag()
    {
        var item = new ListItem
        {
            Id = "42",
            CreatedDateTime = new DateTimeOffset(2026, 03, 25, 9, 0, 0, TimeSpan.Zero),
            LastModifiedDateTime = new DateTimeOffset(2026, 03, 25, 10, 0, 0, TimeSpan.Zero),
            Fields = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object?>
                {
                    ["Title"] = "Client document",
                    ["ClientId"] = "C-123"
                }
            }
        };

        var result = SharePointMapper.ToModel(item);

        Assert.Equal("42", result.Id);
        Assert.Equal("Client document", result.Fields["Title"]);
        Assert.Equal("C-123", result.Fields["ClientId"]);
    }
}
