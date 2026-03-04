using AFH.Common.SharePointUtils.Connector;
using AFH.Common.SharePointUtils.Connector.Interface;
using AFH.Integration.Sharepoint.Models;
using AFH.Common.SharePointUtils.Models;
using AFH.Common.SharePointUtils.Services.Implementation;
using AFH.Common.SharePointUtils.Services.Interface;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;

namespace AFH.Common.SharePointUtils.Configuration;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddSharePoint(
        this IServiceCollection services,
        IConfiguration configuration)
    {
        if (services is null) throw new ArgumentNullException(nameof(services));
        if (configuration is null) throw new ArgumentNullException(nameof(configuration));

        // 1) AzureAD / Graph auth
        var graphConfig = configuration
            .GetSection("AzureAD")
            .Get<SharePointAuthOptions>()
            ?? throw new InvalidOperationException(
                "AzureAD configuration section is missing or invalid.");

        services.Configure<SharePointListsOptions>(
            configuration.GetSection("SharePointListsConfig"));


        services.AddSingleton<GraphServiceClient>(_ =>
        {
            var credential = new ClientSecretCredential(
                graphConfig.TenantId,
                graphConfig.ClientId,
                graphConfig.ClientSecret,
                new TokenCredentialOptions
                {
                    AuthorityHost = new Uri(graphConfig.AuthorityHost)
                });

            return new GraphServiceClient(credential, graphConfig.Scopes);
        });


        services.AddSingleton<ISharePointConnector, SharePointConnector>();
        services.AddSingleton<ISharePointListService, SharePointListService>();
        services.AddSingleton<IClientDocumentService, ClientDocumentService>();

        return services;
    }
}