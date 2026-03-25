using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Tests.Helpers;
using AFH.Common.SharePointUtils.Tests.Fixtures;
using FluentAssertions;
using Xunit.Abstractions;

namespace AFH.Common.SharePointUtils.Tests.Lists;

[Collection("SharepointIntegration")]
[Trait("Area", "Sharepoint")]
[Trait("Kind", "Integration")]
[Trait("Op", "Update")]
public sealed class ListItemUpdateIntegrationTests
{
    private readonly ISharePointListService _svc;
    private readonly ITestOutputHelper _output;

    public ListItemUpdateIntegrationTests(
        SharepointIntegrationFixture fx,
        ITestOutputHelper output)
    {
        _svc = fx.Get<ISharePointListService>();
        _output = output;
    }

    [RequiresSharePointIntegrationFact]
    public async Task UpdateListItem_updates_fields_correctly()
    {
        var siteId = Env("ClientDocuments__SiteId");
        var listId = Env("ClientDocuments__ListId");

        // Arrange: create an initial item
        var initialTitle = $"UpdateTest-{Guid.NewGuid():N}";
        var created = await _svc.AddListItem(
            siteId,
            listId,
            new Dictionary<string, object>
            {
                ["Title"] = initialTitle,
                ["Transcription"] = "BeforeUpdate"
            });

        created.Should().NotBeNull();
        created!.Id.Should().NotBeNullOrWhiteSpace();

        _output.WriteLine($"Created item ID for update test: {created.Id}");

        // Act: update the item
        var updatedTitle = $"{initialTitle}-UPDATED";
        var updated = await _svc.UpdateListItem(
            siteId,
            listId,
            created.Id!,
            new Dictionary<string, object>
            {
                ["Title"] = updatedTitle,
                ["CustomField"] = "AfterUpdate"
            });

        // Assert
        updated.Should().NotBeNull();
        updated!.Fields.Should().NotBeNull();

        var fields = updated.Fields!.AdditionalData;
        fields.Should().NotBeNull();

        fields!["Title"].Should().Be(updatedTitle);
        fields["CustomField"].Should().Be("AfterUpdate");
    }

    private static string Env(string name) =>
        Environment.GetEnvironmentVariable(name)
        ?? throw new InvalidOperationException($"Environment variable '{name}' is not set.");
}