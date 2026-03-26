namespace AFH.Common.SharePointUtils.Configuration;

public sealed class SharePointListConfig
{
    public string SiteId { get; set; } = string.Empty;
    public string ListId { get; set; } = string.Empty;
    public string? ConfigListId { get; set; } = string.Empty;
}
