using AFH.Common.SharePointUtils.Models;

namespace AFH.Common.SharePointUtils.Abstractions;

public interface ISharePointFieldResolver
{
    Task<IReadOnlyList<SharePointFieldDefinition>> GetFieldsAsync(string siteId, string listId, CancellationToken cancellationToken = default);
    Task<SharePointFieldDefinition?> FindFieldAsync(string siteId, string listId, string fieldNameOrDisplayName, CancellationToken cancellationToken = default);
    Task<string?> ResolveInternalNameAsync(string siteId, string listId, string fieldNameOrDisplayName, CancellationToken cancellationToken = default);
    Task<string?> ResolveDisplayNameAsync(string siteId, string listId, string internalName, CancellationToken cancellationToken = default);
    Task<ResolvedSharePointFieldProfile> ResolveProfileAsync(string siteId, string listId, ISharePointMappingProfile profile, CancellationToken cancellationToken = default);
}
