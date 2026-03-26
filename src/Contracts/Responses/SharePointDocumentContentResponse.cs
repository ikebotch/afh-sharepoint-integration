using AFH.Common.SharePointUtils.Models;

namespace AFH.Common.SharePointUtils.Contracts.Responses;

public sealed class SharePointDocumentContentResponse
{
    public required SharePointDocumentMetadata Metadata { get; init; }
    public required Stream Content { get; init; }
}
