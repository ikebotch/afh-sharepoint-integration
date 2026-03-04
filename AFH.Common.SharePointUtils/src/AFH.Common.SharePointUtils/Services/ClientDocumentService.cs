using AFH.Common.SharePointUtils.Connector.Interface;
using AFH.Common.SharePointUtils.Mapping;
using AFH.Common.SharePointUtils.Models;
using AFH.Common.SharePointUtils.Services.Interface;
using Microsoft.Extensions.Logging;

namespace AFH.Common.SharePointUtils.Services.Implementation;

public class ClientDocumentService : IClientDocumentService
{
    private readonly ISharePointConnector _connector;
    private readonly ILogger<ClientDocumentService> _logger;
    private readonly string _siteId;
    private readonly string _listId;

    public ClientDocumentService(
        ISharePointConnector connector,
        ILogger<ClientDocumentService> logger,
        string siteId,
        string listId)
    {
        _connector = connector;
        _logger = logger;
        _siteId = siteId;
        _listId = listId;
    }

    public async Task<bool> VerifyConfigurationAsync(CancellationToken cancellationToken = default)
    {
        var result = await _connector.ListExists(_siteId, _listId, cancellationToken);
        return result.Success && result.Data;
    }

    public async Task<IReadOnlyList<ClientDocument>> GetDocumentsForClientAsync(
        string clientId,
        CancellationToken cancellationToken = default)
    {
        var filter = $"fields/ClientId eq '{clientId}'";
        var result = await _connector.GetListData(_siteId, _listId, filter: filter, cancellationToken: cancellationToken);

        if (!result.Success || result.Data?.Value == null)
            return Array.Empty<ClientDocument>();

        return result.Data.Value.Select(ClientDocumentMapper.ToDomain).ToList();
    }

    public async Task<ClientDocument?> GetDocumentByIdAsync(
        string documentId,
        CancellationToken cancellationToken = default)
    {
        var result = await _connector.GetListItem(_siteId, _listId, documentId, cancellationToken);
        return result.Success && result.Data != null ? ClientDocumentMapper.ToDomain(result.Data) : null;
    }

    public async Task<ClientDocument?> AddDocumentAsync(
        ClientDocument document,
        CancellationToken cancellationToken = default)
    {
        var item = ClientDocumentMapper.ToListItem(document);
        var result = await _connector.AddListItem(_siteId, _listId, item, cancellationToken);
        return result.Success && result.Data != null ? ClientDocumentMapper.ToDomain(result.Data) : null;
    }

    public async Task<ClientDocument?> UpdateDocumentAsync(
        ClientDocument document,
        CancellationToken cancellationToken = default)
    {
        var item = ClientDocumentMapper.ToListItem(document);
        var result = await _connector.UpdateListItem(_siteId, _listId, document.Id, item, cancellationToken);
        return result.Success && result.Data != null ? ClientDocumentMapper.ToDomain(result.Data) : null;
    }

    public async Task<bool> DeleteDocumentAsync(
        string documentId,
        CancellationToken cancellationToken = default)
    {
        var result = await _connector.DeleteListItem(_siteId, _listId, documentId, cancellationToken);
        return result.Success;
    }
}
