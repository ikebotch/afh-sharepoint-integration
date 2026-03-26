using AFH.Common.SharePointUtils.Mapping;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Tests.Mapping;

public class SharePointFieldDefinitionMapperTests
{
    [Fact]
    public void ToModel_Maps_Display_Internal_And_Common_Metadata()
    {
        var column = new ColumnDefinition
        {
            DisplayName = "Client Name",
            Name = "Client_x0020_Name",
            Hidden = false,
            Required = true,
            Text = new TextColumn()
        };

        var result = SharePointFieldDefinitionMapper.ToModel(column);

        Assert.Equal("Client Name", result.DisplayName);
        Assert.Equal("Client_x0020_Name", result.InternalName);
        Assert.Equal("Text", result.Type);
        Assert.True(result.IsRequired);
        Assert.False(result.IsHidden);
    }
}
