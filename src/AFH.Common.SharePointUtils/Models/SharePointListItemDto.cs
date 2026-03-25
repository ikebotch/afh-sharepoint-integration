namespace AFH.Common.SharePointUtils.Models;

public class SharePointListItemDto
{
    public string? Id { get; set; }
    public string? Title { get; set; }
    public DateTimeOffset? Created { get; set; }
    public DateTimeOffset? Modified { get; set; }
}
