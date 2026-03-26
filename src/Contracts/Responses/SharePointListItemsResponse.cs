using AFH.Common.SharePointUtils.Models;

namespace AFH.Common.SharePointUtils.Contracts.Responses;

public sealed class SharePointListItemsResponse
{
    public IReadOnlyList<SharePointListItemModel> Items { get; init; } = [];
    public string? NextPageLink { get; init; }
}
