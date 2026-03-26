using AFH.Common.SharePointUtils.Abstractions;

namespace AFH.Common.SharePointUtils.Models;

public sealed class SharePointMappingProfile : ISharePointMappingProfile
{
    private readonly List<SharePointFieldMap> _fields = [];

    public IReadOnlyCollection<SharePointFieldMap> Fields => _fields;

    public SharePointMappingProfile MapInternal(string propertyName, string internalName)
    {
        _fields.Add(new SharePointFieldMap
        {
            PropertyName = propertyName,
            FieldName = internalName,
            NameKind = SharePointFieldNameKind.InternalName
        });

        return this;
    }

    public SharePointMappingProfile MapDisplay(string propertyName, string displayName)
    {
        _fields.Add(new SharePointFieldMap
        {
            PropertyName = propertyName,
            FieldName = displayName,
            NameKind = SharePointFieldNameKind.DisplayName
        });

        return this;
    }
}
