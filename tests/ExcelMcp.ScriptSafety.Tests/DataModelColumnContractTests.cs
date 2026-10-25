using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.ScriptSafety.Tests;

[Trait("RequiresExcel", "false")]
[Trait("Feature", "PreCommit")]
public sealed class DataModelColumnContractTests
{
    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    [Fact]
    public void UnavailableCalculationFlag_IsExplicitJsonNullInsteadOfFalseOrOmitted()
    {
        var column = new DataModelColumnInfo { Name = "Value" };

        using var json = JsonDocument.Parse(JsonSerializer.Serialize(column, Options));

        Assert.Equal("Value", json.RootElement.GetProperty("name").GetString());
        Assert.Equal(JsonValueKind.Null, json.RootElement.GetProperty("isCalculated").ValueKind);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void KnownCalculationFlag_RetainsItsBooleanValue(bool value)
    {
        var column = new DataModelColumnInfo { Name = "Value", IsCalculated = value };

        using var json = JsonDocument.Parse(JsonSerializer.Serialize(column, Options));

        Assert.Equal(value, json.RootElement.GetProperty("isCalculated").GetBoolean());
    }
}
