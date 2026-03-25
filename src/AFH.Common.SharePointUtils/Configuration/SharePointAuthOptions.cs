using Azure.Identity;

namespace AFH.Common.SharePointUtils.Configuration;

public sealed class SharePointAuthOptions
{
    public required string TenantId { get; set; }
    public required string ClientId { get; set; }
    public required string ClientSecret { get; set; }
    public string[] Scopes { get; set; } = ["https://graph.microsoft.com/.default"];
    public string AuthorityHost { get; set; } = AzureAuthorityHosts.AzurePublicCloud.ToString();
}
