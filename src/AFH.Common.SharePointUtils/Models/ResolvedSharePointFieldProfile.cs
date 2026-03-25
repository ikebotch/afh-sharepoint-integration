namespace AFH.Common.SharePointUtils.Models;

public sealed class ResolvedSharePointFieldProfile
{
    public IReadOnlyDictionary<string, string> PropertyToInternalName { get; init; }
        = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    public bool TryGetInternalName(string propertyName, out string internalName)
        => PropertyToInternalName.TryGetValue(propertyName, out internalName!);

    public string GetRequiredInternalName(string propertyName)
        => PropertyToInternalName.TryGetValue(propertyName, out var internalName)
            ? internalName
            : throw new KeyNotFoundException($"No mapped SharePoint field exists for property '{propertyName}'.");
}
