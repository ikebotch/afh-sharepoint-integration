using AFH.Common.SharePointUtils.Models;
using Microsoft.Graph.Models;

namespace AFH.Common.SharePointUtils.Mapping;

public static class SharePointFieldDefinitionMapper
{
    public static SharePointFieldDefinition ToModel(ColumnDefinition column)
        => new()
        {
            DisplayName = column.DisplayName ?? column.Name ?? string.Empty,
            InternalName = column.Name ?? column.DisplayName ?? string.Empty,
            Type = GetFieldType(column),
            IsRequired = column.Required ?? false,
            IsHidden = column.Hidden ?? false
        };

    private static string GetFieldType(ColumnDefinition column)
    {
        if (column.Text is not null) return "Text";
        if (column.Boolean is not null) return "Boolean";
        if (column.Number is not null) return "Number";
        if (column.Currency is not null) return "Currency";
        if (column.DateTime is not null) return "DateTime";
        if (column.Choice is not null) return "Choice";
        if (column.Lookup is not null) return "Lookup";
        if (column.PersonOrGroup is not null) return "PersonOrGroup";
        if (column.HyperlinkOrPicture is not null) return "HyperlinkOrPicture";
        if (column.Term is not null) return "Term";

        return "Unknown";
    }
}
