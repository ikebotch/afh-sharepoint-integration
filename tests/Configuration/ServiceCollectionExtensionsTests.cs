using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Configuration;
using AFH.Common.SharePointUtils.Extensions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace AFH.Common.SharePointUtils.Tests.Configuration;

public class ServiceCollectionExtensionsTests
{
    [Fact]
    public void AddSharePoint_Registers_GenericClients_And_CompatibilityServices()
    {
        var configuration = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["AzureAD:TenantId"] = "tenant-id",
                ["AzureAD:ClientId"] = "client-id",
                ["AzureAD:ClientSecret"] = "client-secret",
                ["AzureAD:AuthorityHost"] = "https://login.microsoftonline.com/",
                ["AzureAD:Scopes:0"] = "https://graph.microsoft.com/.default",
                ["SharePointListsConfig:ClientDocuments:SiteId"] = "site-id",
                ["SharePointListsConfig:ClientDocuments:ListId"] = "list-id"
            })
            .Build();

        var services = new ServiceCollection();

        services.AddLogging();
        services.AddSharePoint(configuration);

        using var provider = services.BuildServiceProvider();

        Assert.NotNull(provider.GetRequiredService<ISharePointListClient>());
        Assert.NotNull(provider.GetRequiredService<ISharePointDocumentClient>());
        Assert.NotNull(provider.GetRequiredService<ISharePointDiscoveryClient>());
        Assert.NotNull(provider.GetRequiredService<ISharePointFieldResolver>());
        Assert.NotNull(provider.GetRequiredService<ISharePointListService>());
        Assert.NotNull(provider.GetRequiredService<IClientDocumentService>());

        var options = provider.GetRequiredService<Microsoft.Extensions.Options.IOptions<SharePointListsOptions>>().Value;
        Assert.True(options.SharePointListsConfig.ContainsKey("ClientDocuments"));
    }
}
