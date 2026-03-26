using AFH.Common.SharePointUtils.Models;

namespace AFH.Common.SharePointUtils.Abstractions;

public interface ISharePointDiscoveryClient
{
    Task<SharePointSiteReference?> GetSiteAsync(string siteId, CancellationToken cancellationToken = default);
    Task<IReadOnlyList<SharePointListReference>> GetListsAsync(string siteId, CancellationToken cancellationToken = default);
}
