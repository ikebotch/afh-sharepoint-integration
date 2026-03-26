using AFH.Common.SharePointUtils.Extensions;
using AFH.Common.SharePointUtils.Models;
using Microsoft.Kiota.Abstractions.Serialization;

namespace AFH.Common.SharePointUtils.Tests.Extensions;

public class SharePointFieldValueExtensionsTests
{
    [Fact]
    public void Typed_Helpers_Read_Common_Field_Types()
    {
        var fields = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["Title"] = "Client document",
            ["IsActive"] = "true",
            ["Age"] = "42",
            ["Score"] = 98.5,
            ["CreatedOn"] = "2026-03-25T12:30:00Z"
        };

        Assert.Equal("Client document", fields.GetString("title"));
        Assert.True(fields.GetBoolean("IsActive"));
        Assert.Equal(42, fields.GetInt32("Age"));
        Assert.Equal(98.5, fields.GetDouble("Score"));
        Assert.Equal(DateTimeOffset.Parse("2026-03-25T12:30:00Z"), fields.GetDateTimeOffset("CreatedOn"));
    }

    [Fact]
    public void Typed_Helpers_Can_Use_A_Resolved_Field_Profile()
    {
        var profile = new ResolvedSharePointFieldProfile
        {
            PropertyToInternalName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["ClientName"] = "Client_x0020_Name"
            }
        };

        var fields = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["Client_x0020_Name"] = "Jamie Doe"
        };

        Assert.Equal("Jamie Doe", fields.GetString(profile, "ClientName"));
    }

    [Fact]
    public void GetChoices_Reads_Kiota_Untyped_Arrays()
    {
        var fields = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["Skills"] = new UntypedArray(
            [
                new UntypedString("Pension"),
                new UntypedString("Investment")
            ])
        };

        var result = fields.GetChoices("Skills");

        Assert.Equal(["Pension", "Investment"], result);
    }
}
