using AFH.Common.SharePointUtils.Contracts.Requests;
using AFH.Common.SharePointUtils.Contracts.Responses;
using AFH.Common.SharePointUtils.Models;

namespace AFH.Common.SharePointUtils.Abstractions;

public interface ISharePointDocumentClient
{
    Task<IReadOnlyList<SharePointDocumentMetadata>> ListDocumentsAsync(SharePointDocumentQueryRequest request, CancellationToken cancellationToken = default);
    Task<SharePointDocumentMetadata?> GetDocumentMetadataAsync(string siteId, string driveId, string itemId, CancellationToken cancellationToken = default);
    Task<SharePointDocumentContentResponse?> DownloadDocumentAsync(string siteId, string driveId, string itemId, CancellationToken cancellationToken = default);
    Task<SharePointDocumentMetadata?> UploadDocumentAsync(SharePointDocumentUploadRequest request, CancellationToken cancellationToken = default);
    Task<SharePointDocumentMetadata?> UpdateDocumentMetadataAsync(UpdateSharePointDocumentMetadataRequest request, CancellationToken cancellationToken = default);
    Task DeleteDocumentAsync(string siteId, string driveId, string itemId, CancellationToken cancellationToken = default);
}
