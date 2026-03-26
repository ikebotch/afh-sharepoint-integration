using AFH.Common.SharePointUtils.Clients;
using AFH.Common.SharePointUtils.Models;

namespace AFH.Common.SharePointUtils.Tests.Mapping;

public class SharePointFieldResolverTests
{
    [Fact]
    public void FindField_Resolves_By_Display_Or_Internal_Name_Case_Insensitively()
    {
        var fields = CreateFields();

        var byDisplay = GraphSharePointFieldResolver.FindField(fields, "client name");
        var byInternal = GraphSharePointFieldResolver.FindField(fields, "CLIENT_x0020_NAME");

        Assert.NotNull(byDisplay);
        Assert.NotNull(byInternal);
        Assert.Equal("Client_x0020_Name", byDisplay!.InternalName);
        Assert.Equal("Client Name", byInternal!.DisplayName);
    }

    [Fact]
    public void ResolveProfile_Maps_Display_And_Internal_Name_Entries()
    {
        var fields = CreateFields();
        var profile = new SharePointMappingProfile()
            .MapDisplay("ClientName", "Client Name")
            .MapInternal("PolicyNumber", "PolicyNumber");

        var resolved = GraphSharePointFieldResolver.ResolveProfile(fields, profile);

        Assert.Equal("Client_x0020_Name", resolved.GetRequiredInternalName("ClientName"));
        Assert.Equal("PolicyNumber", resolved.GetRequiredInternalName("PolicyNumber"));
    }

    [Fact]
    public async Task GetFieldsAsync_Caches_Field_Definitions_Per_Site_And_List()
    {
        var loadCount = 0;
        var resolver = new GraphSharePointFieldResolver((siteId, listId, _) =>
        {
            loadCount++;
            return Task.FromResult<IReadOnlyList<SharePointFieldDefinition>>(CreateFields());
        });

        var first = await resolver.GetFieldsAsync("site-a", "list-a");
        var second = await resolver.GetFieldsAsync("site-a", "list-a");

        Assert.Equal(1, loadCount);
        Assert.Same(first, second);
    }

    [Fact]
    public void ResolveProfile_Throws_When_A_Field_Cannot_Be_Resolved()
    {
        var profile = new SharePointMappingProfile()
            .MapDisplay("Missing", "Does Not Exist");

        Assert.Throws<KeyNotFoundException>(() =>
            GraphSharePointFieldResolver.ResolveProfile(CreateFields(), profile));
    }

    private static IReadOnlyList<SharePointFieldDefinition> CreateFields()
        =>
        [
            new SharePointFieldDefinition
            {
                DisplayName = "Client Name",
                InternalName = "Client_x0020_Name",
                Type = "Text"
            },
            new SharePointFieldDefinition
            {
                DisplayName = "Policy Number",
                InternalName = "PolicyNumber",
                Type = "Text"
            }
        ];
}
