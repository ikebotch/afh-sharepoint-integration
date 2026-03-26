using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Mapping;
using AFH.Common.SharePointUtils.Models;
using Microsoft.Graph;

namespace AFH.Common.SharePointUtils.Clients;

public sealed class GraphSharePointFieldResolver : ISharePointFieldResolver
{
    private readonly Func<string, string, CancellationToken, Task<IReadOnlyList<SharePointFieldDefinition>>> _fieldLoader;
    private readonly Dictionary<string, IReadOnlyList<SharePointFieldDefinition>> _cache = new(StringComparer.OrdinalIgnoreCase);
    private readonly object _sync = new();

    public GraphSharePointFieldResolver(GraphServiceClient graph)
        : this(async (siteId, listId, cancellationToken) =>
        {
            var response = await graph.Sites[siteId].Lists[listId].Columns.GetAsync(cancellationToken: cancellationToken);
            return response?.Value?.Select(SharePointFieldDefinitionMapper.ToModel).ToList() ?? [];
        })
    {
    }

    public GraphSharePointFieldResolver(
        Func<string, string, CancellationToken, Task<IReadOnlyList<SharePointFieldDefinition>>> fieldLoader)
    {
        _fieldLoader = fieldLoader;
    }

    public async Task<IReadOnlyList<SharePointFieldDefinition>> GetFieldsAsync(string siteId, string listId, CancellationToken cancellationToken = default)
    {
        var cacheKey = GetCacheKey(siteId, listId);

        lock (_sync)
        {
            if (_cache.TryGetValue(cacheKey, out var cached))
                return cached;
        }

        var loaded = await _fieldLoader(siteId, listId, cancellationToken);

        lock (_sync)
        {
            _cache[cacheKey] = loaded;
        }

        return loaded;
    }

    public async Task<SharePointFieldDefinition?> FindFieldAsync(string siteId, string listId, string fieldNameOrDisplayName, CancellationToken cancellationToken = default)
    {
        var fields = await GetFieldsAsync(siteId, listId, cancellationToken);
        return FindField(fields, fieldNameOrDisplayName);
    }

    public async Task<string?> ResolveInternalNameAsync(string siteId, string listId, string fieldNameOrDisplayName, CancellationToken cancellationToken = default)
        => (await FindFieldAsync(siteId, listId, fieldNameOrDisplayName, cancellationToken))?.InternalName;

    public async Task<string?> ResolveDisplayNameAsync(string siteId, string listId, string internalName, CancellationToken cancellationToken = default)
    {
        var fields = await GetFieldsAsync(siteId, listId, cancellationToken);
        return fields.FirstOrDefault(field => string.Equals(field.InternalName, internalName, StringComparison.OrdinalIgnoreCase))?.DisplayName;
    }

    public async Task<ResolvedSharePointFieldProfile> ResolveProfileAsync(string siteId, string listId, ISharePointMappingProfile profile, CancellationToken cancellationToken = default)
    {
        var fields = await GetFieldsAsync(siteId, listId, cancellationToken);
        return ResolveProfile(fields, profile);
    }

    public static SharePointFieldDefinition? FindField(IEnumerable<SharePointFieldDefinition> fields, string fieldNameOrDisplayName)
        => fields.FirstOrDefault(field =>
            string.Equals(field.InternalName, fieldNameOrDisplayName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(field.DisplayName, fieldNameOrDisplayName, StringComparison.OrdinalIgnoreCase));

    public static ResolvedSharePointFieldProfile ResolveProfile(IEnumerable<SharePointFieldDefinition> fields, ISharePointMappingProfile profile)
    {
        var propertyToInternalName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var map in profile.Fields)
        {
            var internalName = map.NameKind switch
            {
                SharePointFieldNameKind.InternalName => map.FieldName,
                SharePointFieldNameKind.DisplayName => FindField(fields, map.FieldName)?.InternalName,
                _ => FindField(fields, map.FieldName)?.InternalName ?? map.FieldName
            };

            if (string.IsNullOrWhiteSpace(internalName))
            {
                throw new KeyNotFoundException(
                    $"Could not resolve SharePoint field '{map.FieldName}' for property '{map.PropertyName}'.");
            }

            propertyToInternalName[map.PropertyName] = internalName;
        }

        return new ResolvedSharePointFieldProfile
        {
            PropertyToInternalName = propertyToInternalName
        };
    }

    private static string GetCacheKey(string siteId, string listId)
        => $"{siteId}::{listId}";
}
