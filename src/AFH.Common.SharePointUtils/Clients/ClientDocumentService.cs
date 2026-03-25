using AFH.Common.SharePointUtils.Abstractions;
using AFH.Common.SharePointUtils.Configuration;
using AFH.Common.SharePointUtils.Contracts.Requests;
using AFH.Common.SharePointUtils.Mapping;
using AFH.Common.SharePointUtils.Models;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace AFH.Common.SharePointUtils.Clients;

public class ClientDocumentService : IClientDocumentService
{
    private readonly ISharePointListClient _listClient;
    private readonly ILogger<ClientDocumentService> _logger;
    private readonly string _siteId;
    private readonly string _listId;

    public ClientDocumentService(
        ISharePointListClient listClient,
        ILogger<ClientDocumentService> logger,
        IOptions<SharePointListsOptions> listOptions)
    {
        _listClient = listClient;
        _logger = logger;

        if (!listOptions.Value.SharePointListsConfig.TryGetValue("ClientDocuments", out var config))
        {
            throw new InvalidOperationException(
                "SharePointListsConfig:ClientDocuments configuration is missing.");
        }

        if (string.IsNullOrWhiteSpace(config.SiteId) || string.IsNullOrWhiteSpace(config.ListId))
        {
            throw new InvalidOperationException(
                "SharePointListsConfig:ClientDocuments must define both SiteId and ListId.");
        }

        _siteId = config.SiteId;
        _listId = config.ListId;
    }

    public Task<bool> VerifyConfigurationAsync(CancellationToken cancellationToken = default)
        => _listClient.ListExistsAsync(_siteId, _listId, cancellationToken);

    public async Task<IReadOnlyList<ClientDocument>> GetDocumentsForClientAsync(
        string clientId,
        CancellationToken cancellationToken = default)
    {
        var result = await _listClient.QueryListItemsAsync(new SharePointListQueryRequest
        {
            SiteId = _siteId,
            ListId = _listId,
            Filter = $"fields/ClientId eq '{clientId}'"
        }, cancellationToken);

        return result.Items.Select(ClientDocumentMapper.ToDomain).ToList();
    }

    public async Task<ClientDocument?> GetDocumentByIdAsync(
        string documentId,
        CancellationToken cancellationToken = default)
    {
        var result = await _listClient.GetListItemAsync(_siteId, _listId, documentId, cancellationToken);
        return result is null ? null : ClientDocumentMapper.ToDomain(result);
    }

    public async Task<ClientDocument?> AddDocumentAsync(
        ClientDocument document,
        CancellationToken cancellationToken = default)
    {
        var result = await _listClient.CreateListItemAsync(new SharePointListItemUpsertRequest
        {
            SiteId = _siteId,
            ListId = _listId,
            Fields = ClientDocumentMapper.ToFields(document)
        }, cancellationToken);

        return result is null ? null : ClientDocumentMapper.ToDomain(result);
    }

    public async Task<ClientDocument?> UpdateDocumentAsync(
        ClientDocument document,
        CancellationToken cancellationToken = default)
    {
        var result = await _listClient.UpdateListItemAsync(new SharePointListItemUpsertRequest
        {
            SiteId = _siteId,
            ListId = _listId,
            ItemId = document.Id,
            Fields = ClientDocumentMapper.ToFields(document)
        }, cancellationToken);

        return result is null ? null : ClientDocumentMapper.ToDomain(result);
    }

    public async Task<bool> DeleteDocumentAsync(
        string documentId,
        CancellationToken cancellationToken = default)
    {
        await _listClient.DeleteListItemAsync(_siteId, _listId, documentId, cancellationToken);
        return true;
    }
}
