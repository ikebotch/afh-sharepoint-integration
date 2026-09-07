using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Clients;
using AFH.Common.SharePointUtils.Configuration;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;

namespace AFH.Common.SharePointUtils.Extensions;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddSharePoint(
        this IServiceCollection services,
        IConfiguration configuration)
    {
        if (services is null) throw new ArgumentNullException(nameof(services));
        if (configuration is null) throw new ArgumentNullException(nameof(configuration));

        var graphConfig = ResolveAuthOptions(configuration);

        services.Configure<SharePointListsOptions>(options =>
        {
            var configuredLists =
                configuration.GetSection("SharePointListsConfig")
                    .Get<Dictionary<string, SharePointListConfig>>();

            if (configuredLists is not null)
            {
                foreach (var (key, value) in configuredLists)
                {
                    options.SharePointListsConfig[key] = value;
                }
            }

            var clientDocumentsSection = configuration.GetSection("ClientDocuments");
            if (clientDocumentsSection.Exists())
            {
                options.SharePointListsConfig["ClientDocuments"] = new SharePointListConfig
                {
                    SiteId = clientDocumentsSection["SiteId"] ?? string.Empty,
                    ListId = clientDocumentsSection["ListId"] ?? string.Empty,
                    ConfigListId = clientDocumentsSection["ConfigListId"] ?? string.Empty
                };
            }
        });

        var advisersSection = configuration.GetSection("SharePoint:Advisers");
        if (advisersSection.Exists())
        {
            services.PostConfigure<SharePointListsOptions>(options =>
            {
                options.SharePointListsConfig["Advisers"] = new SharePointListConfig
                {
                    SiteId = advisersSection["SiteId"] ?? string.Empty,
                    ListId = advisersSection["ListId"] ?? string.Empty
                };
            });
        }

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
        services.AddSingleton<ISharePointListClient, GraphSharePointListClient>();
        services.AddSingleton<ISharePointDocumentClient, GraphSharePointDocumentClient>();
        services.AddSingleton<ISharePointDiscoveryClient, GraphSharePointDiscoveryClient>();
        services.AddSingleton<ISharePointFieldResolver, GraphSharePointFieldResolver>();
        services.AddSingleton<ISharePointListService, SharePointListService>();
        services.AddSingleton<IClientDocumentService, ClientDocumentService>();

        return services;
    }

    internal static SharePointAuthOptions ResolveAuthOptions(IConfiguration configuration)
    {
        var sectionName = configuration.GetSection("SharePointGraph").Exists()
            ? "SharePointGraph"
            : "AzureAD";
        var options = configuration.GetSection(sectionName).Get<SharePointAuthOptions>();

        if (options is null
            || string.IsNullOrWhiteSpace(options.TenantId)
            || string.IsNullOrWhiteSpace(options.ClientId)
            || string.IsNullOrWhiteSpace(options.ClientSecret))
        {
            throw new InvalidOperationException(
                $"{sectionName}:TenantId, {sectionName}:ClientId and {sectionName}:ClientSecret are required for SharePoint Graph authentication.");
        }

        if (!Uri.TryCreate(options.AuthorityHost, UriKind.Absolute, out _))
        {
            throw new InvalidOperationException(
                $"{sectionName}:AuthorityHost must be an absolute URI.");
        }

        if (options.Scopes is not { Length: > 0 }
            || options.Scopes.Any(string.IsNullOrWhiteSpace))
        {
            throw new InvalidOperationException(
                $"{sectionName}:Scopes must contain at least one non-empty scope.");
        }

        return options;
    }
}
