using AFH.Common.SharePointUtils.Services.Interface;
using AFH.Common.SharePointUtils.Tests.Fixtures;
using FluentAssertions;
using Xunit.Abstractions;

namespace AFH.Common.SharePointUtils.Tests.Lists;

[Collection("SharepointIntegration")]
[Trait("Area", "Sharepoint")]
[Trait("Kind", "Integration")]
[Trait("Op", "Delete")]
public sealed class ListItemDeleteIntegrationTests
{
    private readonly ISharePointListService _svc;
    private readonly ITestOutputHelper _output;

    public ListItemDeleteIntegrationTests(
        SharepointIntegrationFixture fx,
        ITestOutputHelper output)
    {
        _svc = fx.Get<ISharePointListService>();
        _output = output;
    }

    [Fact]
    public async Task DeleteListItem_removes_item()
    {
        var siteId = Env("ClientDocuments__SiteId");
        var listId = Env("ClientDocuments__ListId");

        // Arrange: create an item to delete
        var title = $"DeleteTest-{Guid.NewGuid():N}";
        var created = await _svc.AddListItem(
            siteId,
            listId,
            new Dictionary<string, object>
            {
                ["Title"] = title,
                ["Transcription"] = "ToBeDeleted"
            });

        created.Should().NotBeNull();
        created!.Id.Should().NotBeNullOrWhiteSpace();

        _output.WriteLine($"Created item ID for delete test: {created.Id}");

        // Act: delete the item
        var deleted = await _svc.DeleteListItem(
            siteId,
            listId,
            created.Id!);

        // Assert: delete call itself should succeed
        deleted.Should().BeTrue("the item should be deleted successfully");

        // And confirm it no longer appears in queries
        var itemsAfterDelete = await _svc.GetListItems(
            siteId,
            listId,
            filter: $"Title eq '{title}'");

        itemsAfterDelete.Should().BeEmpty("the deleted item should no longer exist");
    }

    private static string Env(string name) =>
        Environment.GetEnvironmentVariable(name)
        ?? throw new InvalidOperationException($"Environment variable '{name}' is not set.");
}