namespace AFH.Common.SharePointUtils.Configuration;

public class SharePointListsOptions
{
    public Dictionary<string, SharePointListConfig> SharePointListsConfig { get; set; }
        = new(StringComparer.OrdinalIgnoreCase);
}
