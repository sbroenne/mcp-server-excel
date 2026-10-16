using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceConnectionTests
{
    [ConfiguredSsasDateTheory]
    [InlineData("SpecificDate")]
    [InlineData("NotSpecificDate")]
    [InlineData("Before")]
    [InlineData("BeforeOrEqualTo")]
    [InlineData("After")]
    [InlineData("AfterOrEqualTo")]
    [InlineData("DateBetween")]
    [InlineData("DateNotBetween")]
    [Trait("RunType", "OnDemand")]
    public void SsasDateFilters_TypedDateLevel_FiltersAndClearsThroughService(string filterType)
    {
        bool interval = filterType.Contains("Between", StringComparison.Ordinal);
        string[] expectedDates = filterType switch
        {
            "SpecificDate" => ["2024-02-10"],
            "NotSpecificDate" => ["2024-01-15", "2024-03-05", "2024-04-12"],
            "Before" => ["2024-01-15"],
            "BeforeOrEqualTo" => ["2024-01-15", "2024-02-10"],
            "After" => ["2024-03-05", "2024-04-12"],
            "AfterOrEqualTo" => ["2024-02-10", "2024-03-05", "2024-04-12"],
            "DateBetween" => ["2024-02-10", "2024-03-05"],
            "DateNotBetween" => ["2024-01-15", "2024-04-12"],
            _ => throw new ArgumentOutOfRangeException(nameof(filterType))
        };
        var amounts = new Dictionary<string, double>(StringComparer.Ordinal)
        {
            ["2024-01-15"] = 100,
            ["2024-02-10"] = 200,
            ["2024-03-05"] = 400,
            ["2024-04-12"] = 800
        };
        var connection = UniqueConnectionName("SsasDates");
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        CreateOlapPivotTable(
            Environment.GetEnvironmentVariable("EXCELMCP_TEST_SSAS_DATE_CONNECTION_STRING")!,
            "Model", connection, sheet);
        var name = connection + "_Pivot";
        var commands = _fixture.CreateCommands<IPersistentPivotTableCommands>();
        RequireSuccess(commands.AddRowField(_fixture.BatchToken, name, "[Dates].[Date]", null));
        RequireSuccess(commands.AddValueField(_fixture.BatchToken, name, "[Measures].[Total Amount]"));
        AssertDates(["2024-01-15", "2024-02-10", "2024-03-05", "2024-04-12"], [100, 200, 400, 800]);

        var added = _fixture.Send("pivottablefield.add-field-filter", new
        {
            pivotTableName = name,
            fieldName = "[Dates].[Date].[Date]",
            filterOptions = new
            {
                type = filterType,
                date1 = new DateTime(2024, 2, 10),
                date2 = interval ? new DateTime(2024, 3, 5) : (DateTime?)null
            }
        });
        Assert.True(added.Success, added.ErrorMessage);
        using (var result = JsonDocument.Parse(added.Result!))
            AssertCriteria(result.RootElement);
        AssertDates(expectedDates, expectedDates.Select(date => amounts[date]).ToArray());

        var read = _fixture.Send("pivottablefield.get-field-filters", new
        {
            pivotTableName = name,
            fieldName = "[Dates].[Date]"
        });
        Assert.True(read.Success, read.ErrorMessage);
        using (var result = JsonDocument.Parse(read.Result!))
            AssertCriteria(result.RootElement);
        var snapshot = RequireSuccess(commands.Read(_fixture.BatchToken, name, sheet));
        Assert.True(snapshot.SnapshotComplete);
        var selection = Assert.Single(snapshot.FieldSelections);
        Assert.True(selection.FiltersAvailable);
        var criterion = Assert.Single(selection.Filters);
        Assert.Equal(filterType, criterion.Type);
        Assert.Equal("[Dates].[Date].[Date]", criterion.LevelUniqueName);
        Assert.Equal("2024-02-10T00:00:00.0000000", criterion.Value1);
        Assert.Equal(interval ? "2024-03-05T00:00:00.0000000" : null, criterion.Value2);

        var cleared = _fixture.Send("pivottablefield.clear-field-filters", new
        {
            pivotTableName = name,
            fieldName = "[Dates].[Date].[Date]"
        });
        Assert.True(cleared.Success, cleared.ErrorMessage);
        using (var result = JsonDocument.Parse(cleared.Result!))
            Assert.Empty(result.RootElement.GetProperty("filters").EnumerateArray());
        AssertDates(["2024-01-15", "2024-02-10", "2024-03-05", "2024-04-12"], [100, 200, 400, 800]);

        void AssertCriteria(JsonElement result)
        {
            Assert.True(result.GetProperty("success").GetBoolean());
            Assert.Equal("[Dates].[Date].[Date]", result.GetProperty("fieldName").GetString());
            var filter = Assert.Single(result.GetProperty("filters").EnumerateArray());
            Assert.Equal(filterType, filter.GetProperty("type").GetString());
            Assert.Equal("2024-02-10T00:00:00.0000000", filter.GetProperty("value1").GetString());
            if (interval)
                Assert.Equal("2024-03-05T00:00:00.0000000", filter.GetProperty("value2").GetString());
            else
                Assert.True(!filter.TryGetProperty("value2", out var second) || second.ValueKind == JsonValueKind.Null);
        }

        void AssertDates(string[] dates, double[] amounts)
        {
            var data = RequireSuccess(commands.GetData(_fixture.BatchToken, name));
            Assert.Equal(dates.Length + 2, data.Values.Count);
            var rows = data.Values.Skip(1).Take(dates.Length).ToArray();
            Assert.Equal(dates, rows.Select(row => Convert.ToString(row[0], System.Globalization.CultureInfo.InvariantCulture)));
            Assert.Equal(amounts, rows.Select(row => Convert.ToDouble(row[1], System.Globalization.CultureInfo.InvariantCulture)));
            Assert.Equal(amounts.Sum(), Convert.ToDouble(data.Values[^1][1], System.Globalization.CultureInfo.InvariantCulture));
        }
    }
}

[AttributeUsage(AttributeTargets.Method)]
public sealed class ConfiguredSsasDateTheoryAttribute : TheoryAttribute
{
    public ConfiguredSsasDateTheoryAttribute()
    {
        if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("EXCELMCP_TEST_SSAS_DATE_CONNECTION_STRING")))
            Skip = "Configure the synthetic SSAS Dates table and Total Amount measure to run this regression.";
    }
}
