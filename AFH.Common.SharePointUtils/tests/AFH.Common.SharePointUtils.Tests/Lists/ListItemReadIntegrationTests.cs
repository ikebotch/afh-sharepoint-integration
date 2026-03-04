using AFH.Common.SharePointUtils.Services.Interface;
using AFH.Common.SharePointUtils.Tests.Fixtures;
using FluentAssertions;
using Xunit.Abstractions;

namespace AFH.Common.SharePointUtils.Tests.Lists;

[Collection("SharepointIntegration")]
[Trait("Area", "Sharepoint")]
[Trait("Kind", "Integration")]
[Trait("Op", "Read")]
public sealed class ListItemReadIntegrationTests
{
    private readonly ISharePointListService _svc;
    private readonly ITestOutputHelper _output;

    public ListItemReadIntegrationTests(
        SharepointIntegrationFixture fx,
        ITestOutputHelper output)
    {
        _svc = fx.Get<ISharePointListService>();
        _output = output;
    }

    [Fact]
    public async Task GetListItems_returns_items_or_empty_but_does_not_fail()
    {
        var siteId = Env("ClientDocuments__SiteId");
        var listId = Env("ClientDocuments__ListId");

        var items = await _svc.GetListItems(
            siteId,
            listId,
            filter: null);

        items.Should().NotBeNull("the call to SharePoint should succeed even if the list is empty");

        _output.WriteLine($"Read {items.Count} items from list '{listId}' on site '{siteId}'.");
    }

    private static string Env(string name) =>
        Environment.GetEnvironmentVariable(name)
        ?? throw new InvalidOperationException($"Environment variable '{name}' is not set.");
}