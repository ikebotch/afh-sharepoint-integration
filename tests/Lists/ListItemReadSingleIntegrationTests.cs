using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Tests.Helpers;
using AFH.Common.SharePointUtils.Tests.Fixtures;
using FluentAssertions;

namespace AFH.Common.SharePointUtils.Tests.Lists;

[Collection("SharepointIntegration")]
[Trait("Area", "Sharepoint")]
[Trait("Kind", "Integration")]
[Trait("Op", "ReadSingle")]
public sealed class ListItemReadSingleIntegrationTests
{
    private readonly ISharePointListService _svc;

    public ListItemReadSingleIntegrationTests(SharepointIntegrationFixture fx)
    {
        _svc = fx.Get<ISharePointListService>();
    }

    [RequiresSharePointIntegrationFact]
    public async Task GetSingleItem_returns_item()
    {
        var siteId = Environment.GetEnvironmentVariable("ClientDocuments__SiteId");
        var listId = Environment.GetEnvironmentVariable("ClientDocuments__ListId");

        var items = await _svc.GetListItems(siteId, listId);
        var first = items.FirstOrDefault();

        first.Should().NotBeNull();
        first.Id.Should().NotBeNullOrWhiteSpace();
    }
}
