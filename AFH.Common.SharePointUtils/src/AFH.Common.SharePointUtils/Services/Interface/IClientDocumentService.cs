using AFH.Common.SharePointUtils.Models;

namespace AFH.Common.SharePointUtils.Services.Interface;

public interface IClientDocumentService
{
    Task<bool> VerifyConfigurationAsync(CancellationToken cancellationToken = default);

    Task<IReadOnlyList<ClientDocument>> GetDocumentsForClientAsync(
        string clientId,
        CancellationToken cancellationToken = default);

    Task<ClientDocument?> GetDocumentByIdAsync(
        string documentId,
        CancellationToken cancellationToken = default);

    Task<ClientDocument?> AddDocumentAsync(
        ClientDocument document,
        CancellationToken cancellationToken = default);

    Task<ClientDocument?> UpdateDocumentAsync(
        ClientDocument document,
        CancellationToken cancellationToken = default);

    Task<bool> DeleteDocumentAsync(
        string documentId,
        CancellationToken cancellationToken = default);
}
