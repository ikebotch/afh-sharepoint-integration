using Azure.Identity;

namespace AFH.Integration.Sharepoint.Models
{
    public class SharePointAuthOptions
    {
        public required string TenantId { get; set; }
        public required string ClientId { get; set; }
        public required string ClientSecret { get; set; }
        public string[] Scopes { get; set; } = new[] { "https://graph.microsoft.com/.default" };
        public string AuthorityHost { get; set; } = AzureAuthorityHosts.AzurePublicCloud.ToString();
    }
}
