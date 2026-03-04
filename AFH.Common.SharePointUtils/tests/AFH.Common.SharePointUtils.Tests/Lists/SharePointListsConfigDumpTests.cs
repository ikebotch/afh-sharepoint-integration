using AFH.Common.SharePointUtils.Models;
using AFH.Common.SharePointUtils.Tests.Fixtures;
using FluentAssertions;
using Microsoft.Extensions.Options;
using Xunit.Abstractions;

namespace AFH.Common.SharePointUtils.Tests.Lists;

[Collection("SharepointIntegration")]
[Trait("Area", "Sharepoint")]
[Trait("Kind", "Integration")]
[Trait("Op", "ConfigDump")]
public sealed class SharePointListsConfigDumpTests
{
    private readonly SharepointIntegrationFixture _fx;
    private readonly ITestOutputHelper _output;

    public SharePointListsConfigDumpTests(
        SharepointIntegrationFixture fx,
        ITestOutputHelper output)
    {
        _fx = fx;
        _output = output;
    }

    [Fact]
    public void Dump_SharePointListsConfig()
    {
        var opts = _fx.Get<IOptions<SharePointListsOptions>>();
        var allConfigs = opts.Value.SharePointListsConfig;

        allConfigs.Should().NotBeNull();

        _output.WriteLine($"Count = {allConfigs.Count} {opts.ToString()}");

        if (!allConfigs.Any())
        {
            _output.WriteLine("No SharePointListsConfig entries were found.");
            return;
        }

        foreach (var (key, cfg) in allConfigs)
        {
            _output.WriteLine($"Key: {key}");
            _output.WriteLine($"  SiteId:       {cfg.SiteId}");
            _output.WriteLine($"  ListId:       {cfg.ListId}");
            _output.WriteLine($"  ConfigList:   {cfg.ConfigListId ?? "(null)"}");
            _output.WriteLine("-------------------------------------------");
        }
    }
}

