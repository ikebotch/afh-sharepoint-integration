using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Models;
using Microsoft.Graph;

namespace AFH.Common.SharePointUtils.Clients;

public sealed class GraphSharePointDiscoveryClient : ISharePointDiscoveryClient
{
    private readonly GraphServiceClient _graph;

    public GraphSharePointDiscoveryClient(GraphServiceClient graph)
    {
        _graph = graph;
    }

    public async Task<SharePointSiteReference?> GetSiteAsync(string siteId, CancellationToken cancellationToken = default)
    {
        var site = await _graph.Sites[siteId].GetAsync(cancellationToken: cancellationToken);
        return site is null
            ? null
            : new SharePointSiteReference
            {
                Id = site.Id,
                Name = site.Name,
                DisplayName = site.DisplayName,
                WebUrl = site.WebUrl
            };
    }

    public async Task<IReadOnlyList<SharePointListReference>> GetListsAsync(string siteId, CancellationToken cancellationToken = default)
    {
        var lists = await _graph.Sites[siteId].Lists.GetAsync(cancellationToken: cancellationToken);
        return lists?.Value?.Select(list => new SharePointListReference
        {
            Id = list.Id,
            Name = list.Name,
            DisplayName = list.DisplayName
        }).ToList() ?? [];
    }
}
