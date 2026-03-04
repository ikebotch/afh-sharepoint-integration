namespace AFH.Common.SharePointUtils.Models;

public class SharePointListsOptions
{

    public Dictionary<string, SharePointListConfig> SharePointListsConfig { get; set; }
        = new(StringComparer.OrdinalIgnoreCase);
}