using Xunit;

namespace AFH.Common.SharePointUtils.Tests.Helpers;

public sealed class RequiresSharePointIntegrationFactAttribute : FactAttribute
{
    public RequiresSharePointIntegrationFactAttribute()
    {
        var requiredVariables = new[]
        {
            "ClientDocuments__SiteId",
            "ClientDocuments__ListId",
            "AzureAD__TenantId",
            "AzureAD__ClientId",
            "AzureAD__ClientSecret"
        };

        if (requiredVariables.Any(name => string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(name))))
        {
            Skip = "Requires SharePoint integration environment variables for a live tenant-backed test run.";
        }
    }
}
