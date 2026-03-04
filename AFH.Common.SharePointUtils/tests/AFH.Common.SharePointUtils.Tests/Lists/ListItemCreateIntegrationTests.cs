using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using AFH.Common.SharePointUtils.Services.Interface;
using AFH.Common.SharePointUtils.Tests.Fixtures;
using FluentAssertions;
using Microsoft.Extensions.Configuration;
using Xunit;
using Xunit.Abstractions;

namespace AFH.Common.SharePointUtils.Tests.Lists;

[Collection("SharepointIntegration")]
[Trait("Area", "Sharepoint")]
[Trait("Kind", "Integration")]
public sealed class ListItemCreateIntegrationTests
{
    private readonly ISharePointListService _svc;
    private readonly IConfiguration _cfg;
    private readonly ITestOutputHelper _output;

    public ListItemCreateIntegrationTests(
        SharepointIntegrationFixture fx,
        ITestOutputHelper output)
    {
        _svc = fx.Get<ISharePointListService>();
        _cfg = fx.Get<IConfiguration>();
        _output = output;
    }

    [Fact]
    [Trait("Op", "Create")]
    public async Task AddListItem_creates_item()
    {
        // Read site/list IDs directly from IConfiguration
        var siteId = _cfg["SharePointListsConfig:ClientDocuments:SiteId"];
        var listId = _cfg["SharePointListsConfig:ClientDocuments:ListId"];

        siteId.Should().NotBeNullOrWhiteSpace("ClientDocuments SiteId must be configured");
        listId.Should().NotBeNullOrWhiteSpace("ClientDocuments ListId must be configured");

        _output.WriteLine($"Using SiteId={siteId}, ListId={listId}");

        // 2Construct test item
        var item = new Dictionary<string, object>
        {
            ["Title"] = $"SDK Integration {Guid.NewGuid()}",
            ["Transcription"] = "Test transcription"
        };

        // Act
        var created = await _svc.AddListItem(siteId!, listId!, item);

        //Assert
        created.Should().NotBeNull("SharePoint should return the created list item");
        created!.Id.Should().NotBeNullOrWhiteSpace("created items must have a valid Id");

        _output.WriteLine(
            $"Created item Id={created.Id} in list {listId} on site {siteId}");
    }
}