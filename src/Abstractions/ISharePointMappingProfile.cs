using AFH.Common.SharePointUtils.Models;

namespace AFH.Common.SharePointUtils.Abstractions;

public interface ISharePointMappingProfile
{
    IReadOnlyCollection<SharePointFieldMap> Fields { get; }
}
