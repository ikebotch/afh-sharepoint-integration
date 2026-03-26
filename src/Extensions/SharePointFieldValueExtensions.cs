using AFH.Common.SharePointUtils.Models;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Serialization;

namespace AFH.Common.SharePointUtils.Extensions;

public static class SharePointFieldValueExtensions
{
    public static IReadOnlyDictionary<string, object?> ToFieldDictionary(this FieldValueSet? fields)
        => fields?.AdditionalData?.ToDictionary(
            kvp => kvp.Key,
            kvp => (object?)kvp.Value,
            StringComparer.OrdinalIgnoreCase)
        ?? new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

    public static IReadOnlyDictionary<string, object?> GetFieldValues(this ListItem? item)
        => item?.Fields.ToFieldDictionary()
        ?? new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

    public static string? GetString(this IReadOnlyDictionary<string, object?> fields, string internalName)
        => fields.TryGetValue(internalName, out var value) ? value?.ToString() : null;

    public static string? GetString(this IReadOnlyDictionary<string, object?> fields, ResolvedSharePointFieldProfile profile, string propertyName)
        => fields.GetString(profile.GetRequiredInternalName(propertyName));

    public static bool? GetBoolean(this IReadOnlyDictionary<string, object?> fields, string internalName)
    {
        if (!fields.TryGetValue(internalName, out var value) || value is null)
            return null;

        if (value is bool boolean)
            return boolean;

        if (bool.TryParse(value.ToString(), out var parsed))
            return parsed;

        return null;
    }

    public static int? GetInt32(this IReadOnlyDictionary<string, object?> fields, string internalName)
    {
        if (!fields.TryGetValue(internalName, out var value) || value is null)
            return null;

        if (value is int intValue)
            return intValue;

        if (int.TryParse(value.ToString(), out var parsed))
            return parsed;

        return null;
    }

    public static double? GetDouble(this IReadOnlyDictionary<string, object?> fields, string internalName)
    {
        if (!fields.TryGetValue(internalName, out var value) || value is null)
            return null;

        if (value is double doubleValue)
            return doubleValue;

        if (double.TryParse(value.ToString(), out var parsed))
            return parsed;

        return null;
    }

    public static DateTimeOffset? GetDateTimeOffset(this IReadOnlyDictionary<string, object?> fields, string internalName)
    {
        if (!fields.TryGetValue(internalName, out var value) || value is null)
            return null;

        if (value is DateTimeOffset dto)
            return dto;

        if (value is DateTime dt)
            return new DateTimeOffset(dt);

        if (DateTimeOffset.TryParse(value.ToString(), out var parsed))
            return parsed;

        return null;
    }

    public static string? GetChoice(this IReadOnlyDictionary<string, object?> fields, string internalName)
        => fields.GetString(internalName);

    public static IReadOnlyList<string> GetChoices(this IReadOnlyDictionary<string, object?> fields, string internalName)
    {
        if (!fields.TryGetValue(internalName, out var value) || value is null)
            return [];

        if (value is IEnumerable<string> strings)
            return strings.Where(value => !string.IsNullOrWhiteSpace(value)).ToList();

        if (value is UntypedArray array)
        {
            return array.GetValue()?
                .OfType<UntypedString>()
                .Select(item => item.GetValue())
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Cast<string>()
                .ToList()
                ?? [];
        }

        return [value.ToString()!];
    }

    public static int? GetLookupId(this IReadOnlyDictionary<string, object?> fields, string internalName)
        => fields.GetInt32(internalName);

    public static string? GetLookupValue(this IReadOnlyDictionary<string, object?> fields, string internalName)
        => fields.GetString(internalName);

    public static string? GetPersonDisplayName(this IReadOnlyDictionary<string, object?> fields, string internalName)
        => fields.GetString(internalName);
}
