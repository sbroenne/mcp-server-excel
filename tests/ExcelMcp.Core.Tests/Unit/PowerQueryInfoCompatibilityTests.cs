using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class PowerQueryInfoCompatibilityTests
{
    [Fact]
    public void Formula_SourceContractRemainsReadWrite()
    {
#pragma warning disable CS0618
        var query = new PowerQueryInfo
        {
            Formula = "let Source = 1 in Source"
        };

        Assert.Equal("let Source = 1 in Source", query.Formula);
#pragma warning restore CS0618
    }

    [Fact]
    public void Formula_RemainsPublicReadWriteAndObsoleteForBinaryCompatibility()
    {
        var property = typeof(PowerQueryInfo).GetProperty(
            "Formula",
            BindingFlags.Instance | BindingFlags.Public);

        Assert.NotNull(property);
        Assert.Equal(typeof(string), property.PropertyType);
        Assert.True(property.CanRead);
        Assert.True(property.CanWrite);
        Assert.NotNull(property.GetMethod);
        Assert.NotNull(property.SetMethod);
        Assert.NotNull(property.GetCustomAttribute<ObsoleteAttribute>());
        Assert.NotNull(property.GetCustomAttribute<JsonIgnoreAttribute>());

        var query = new PowerQueryInfo();
        property.SetValue(query, "let Source = 1 in Source");
        Assert.Equal("let Source = 1 in Source", property.GetValue(query));
    }

    [Fact]
    public void Formula_IsOmittedFromListJsonEvenWhenPopulated()
    {
        var property = typeof(PowerQueryInfo).GetProperty("Formula");
        Assert.NotNull(property);
        var query = new PowerQueryInfo
        {
            Name = "Compatibility",
            FormulaPreview = "let Source = 1 in Source",
            CharacterCount = 24
        };
        property.SetValue(query, "let Source = 1 in Source");

        var json = JsonSerializer.Serialize(
            new PowerQueryListResult
            {
                Success = true,
                Queries = [query]
            },
            JsonSerializerOptions.Web);

        using var document = JsonDocument.Parse(json);
        var serializedQuery = Assert.Single(
            document.RootElement.GetProperty("queries").EnumerateArray());
        Assert.False(serializedQuery.TryGetProperty("formula", out _));
        Assert.Equal(
            "let Source = 1 in Source",
            serializedQuery.GetProperty("formulaPreview").GetString());
    }
}
