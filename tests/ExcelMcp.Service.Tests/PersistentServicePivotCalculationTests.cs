using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PivotCalculation")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePivotCalculationTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPersistentPivotTableCommands _pivot =
        ServiceCommandProxy.Create<IPersistentPivotTableCommands>(fixture);

    [Fact]
    public void PercentOfTotal_ChangesDisplayedValuesAndResetsWithoutChangingAggregation()
    {
        var name = CreatePivot();
        using (var set = Set(name, "PercentOfTotal"))
            Assert.Equal("PercentOfTotal", set.RootElement.GetProperty("calculation").GetString());
        var data = _pivot.GetData(_fixture.BatchToken, name);
        Assert.True(data.Success, data.ErrorMessage);
        AssertRegionRows(data.Values);
        Assert.NotNull(data.Values[1][1]);
        Assert.NotNull(data.Values[2][1]);
        Assert.Equal(0.25d, Convert.ToDouble(data.Values[1][1], CultureInfo.InvariantCulture));
        Assert.Equal(0.75d, Convert.ToDouble(data.Values[2][1], CultureInfo.InvariantCulture));
        using (var reset = Set(name, "Normal"))
            Assert.Equal("Normal", reset.RootElement.GetProperty("calculation").GetString());
        var normal = _pivot.GetData(_fixture.BatchToken, name);
        Assert.True(normal.Success, normal.ErrorMessage);
        AssertRegionRows(normal.Values);
        Assert.NotNull(normal.Values[1][1]);
        Assert.NotNull(normal.Values[2][1]);
        Assert.Equal(100d, Convert.ToDouble(normal.Values[1][1], CultureInfo.InvariantCulture));
        Assert.Equal(300d, Convert.ToDouble(normal.Values[2][1], CultureInfo.InvariantCulture));
        var listed = _fixture.Send("pivottablefield.list-fields", new { pivotTableName = name });
        using var fields = JsonDocument.Parse(listed.Result!);
        var value = Assert.Single(fields.RootElement.GetProperty("valueFields").EnumerateArray());
        Assert.Equal("Normal", value.GetProperty("calculation").GetString());
        Assert.Equal("Sum", value.GetProperty("function").GetString());
    }

    [Fact]
    public void RunningTotal_UsesExplicitBaseFieldAndReturnsNativeResults()
    {
        var name = CreatePivot();
        using var set = Set(name, "RunningTotal", "Region");
        Assert.Equal("Region", set.RootElement.GetProperty("baseFieldName").GetString());
        var data = _pivot.GetData(_fixture.BatchToken, name);
        Assert.True(data.Success, data.ErrorMessage);
        AssertRegionRows(data.Values);
        Assert.NotNull(data.Values[1][1]);
        Assert.NotNull(data.Values[2][1]);
        Assert.Equal(100d, Convert.ToDouble(data.Values[1][1], CultureInfo.InvariantCulture));
        Assert.Equal(400d, Convert.ToDouble(data.Values[2][1], CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData("Normal", null, 100d, 300d)]
    [InlineData("PercentOfRow", null, 1d, 1d)]
    [InlineData("PercentOfColumn", null, 0.25d, 0.75d)]
    [InlineData("Index", null, 1d, 1d)]
    [InlineData("PercentRunningTotal", "Region", 0.25d, 1d)]
    [InlineData("RankAscending", "Region", 1d, 2d)]
    [InlineData("RankDescending", "Region", 2d, 1d)]
    [InlineData("PercentOfParentRow", null, 0.25d, 0.75d)]
    public void Calculations_ReturnActualNativeNumbers(string calculation, string? baseField, double north, double south)
    {
        var name = CreatePivot();
        using var set = Set(name, calculation, baseField);
        Assert.Equal(calculation, set.RootElement.GetProperty("calculation").GetString());
        AssertValues(name, north, south);
    }

    [Theory]
    [InlineData("DifferenceFrom", null, 200d)]
    [InlineData("PercentOf", 1d, 3d)]
    [InlineData("PercentDifferenceFrom", null, 2d)]
    public void NamedBaseItem_PreservesIdentityAndNativeNumbers(string calculation, double? north, double south)
    {
        var name = CreatePivot();
        using var set = Set(name, calculation, "Region", "Named", "North");
        Assert.Equal("Named", set.RootElement.GetProperty("baseItemKind").GetString());
        Assert.Equal("North", set.RootElement.GetProperty("baseItemName").GetString());
        var data = _pivot.GetData(_fixture.BatchToken, name);
        Assert.True(data.Success, data.ErrorMessage);
        AssertRegionRows(data.Values);
        if (north.HasValue)
        {
            Assert.NotNull(data.Values[1][1]);
            Assert.Equal(north.Value, Convert.ToDouble(data.Values[1][1], CultureInfo.InvariantCulture), precision: 8);
        }
        else
        {
            Assert.Null(data.Values[1][1]);
        }
        Assert.NotNull(data.Values[2][1]);
        Assert.Equal(south, Convert.ToDouble(data.Values[2][1], CultureInfo.InvariantCulture), precision: 8);
    }

    [Theory]
    [InlineData("PercentOfParentRow", false, null)]
    [InlineData("PercentOfParent", false, "Region")]
    [InlineData("PercentOfParentColumn", true, null)]
    public void ParentPercentages_UseActualRowOrColumnHierarchy(string calculation, bool columns, string? baseFieldName)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Parent_{Guid.NewGuid():N}";
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:C5",
            [["Region", "Product", "Sales"], ["North", "A", 100], ["North", "B", 300],
                ["South", "A", 200], ["South", "B", 400]]).Success);
        Assert.True(_pivot.CreateFromRange(_fixture.BatchToken, sheet, "A1:C5", sheet, "E1", name).Success);
        foreach (var field in new[] { "Region", "Product" })
        {
            var added = columns
                ? _pivot.AddColumnField(_fixture.BatchToken, name, field)
                : _pivot.AddRowField(_fixture.BatchToken, name, field);
            Assert.True(added.Success, added.ErrorMessage);
            Assert.True(_pivot.SetSubtotals(_fixture.BatchToken, name, field, false).Success);
        }
        Assert.True(_pivot.AddValueField(_fixture.BatchToken, name, "Sales", AggregationFunction.Sum, "Total Sales").Success);
        Assert.True(_pivot.SetLayout(_fixture.BatchToken, name, 1).Success);
        using var set = Set(name, calculation, baseFieldName);
        Assert.Equal(calculation, set.RootElement.GetProperty("calculation").GetString());
        var data = _pivot.GetData(_fixture.BatchToken, name);
        Assert.True(data.Success, data.ErrorMessage);
        double[] expected = [0.25d, 0.75d, 1d / 3d, 2d / 3d];
        for (int i = 0; i < expected.Length; i++)
        {
            var value = columns ? data.Values[^1][i + 1] : data.Values[i + 1][2];
            Assert.NotNull(value);
            Assert.Equal(expected[i], Convert.ToDouble(value, CultureInfo.InvariantCulture), precision: 8);
        }
    }

    [Theory]
    [InlineData("Previous", 2, 200d)]
    [InlineData("Next", 1, -200d)]
    public void RelativeBaseItem_ReadsActualPreviousOrNextSetting(string kind, int valueRow, double expected)
    {
        var name = CreatePivot();
        using var set = Set(name, "DifferenceFrom", "Region", kind);
        Assert.Equal(kind, set.RootElement.GetProperty("baseItemKind").GetString());
        Assert.False(set.RootElement.TryGetProperty("baseItemName", out _));
        var data = _pivot.GetData(_fixture.BatchToken, name);
        Assert.True(data.Success, data.ErrorMessage);
        AssertRegionRows(data.Values);
        Assert.NotNull(data.Values[valueRow][1]);
        Assert.Equal(expected, Convert.ToDouble(data.Values[valueRow][1], CultureInfo.InvariantCulture));
    }

    [Fact]
    public void RepeatedSourceFields_ListEveryInstanceAndChangeOnlySelectedCaption()
    {
        var name = CreatePivot();
        var extra = _pivot.AddValueField(_fixture.BatchToken, name, "Sales", AggregationFunction.Average, "Average Sales");
        Assert.True(extra.Success, extra.ErrorMessage);
        using (var set = Set(name, "PercentOfTotal"))
            Assert.Equal("Total Sales", set.RootElement.GetProperty("fieldName").GetString());
        var listed = _fixture.Send("pivottablefield.list-fields", new { pivotTableName = name });
        using var result = JsonDocument.Parse(listed.Result!);
        var fields = result.RootElement.GetProperty("valueFields").EnumerateArray().ToArray();
        Assert.Equal(2, fields.Length);
        var total = Assert.Single(fields, field => field.GetProperty("fieldName").GetString() == "Total Sales");
        var average = Assert.Single(fields, field => field.GetProperty("fieldName").GetString() == "Average Sales");
        Assert.Equal("PercentOfTotal", total.GetProperty("calculation").GetString());
        Assert.Equal("Normal", average.GetProperty("calculation").GetString());
        Assert.Equal("Average", average.GetProperty("function").GetString());
        Assert.All(fields, field => Assert.Equal("Sales", field.GetProperty("sourceName").GetString()));
        var data = _pivot.GetData(_fixture.BatchToken, name);
        RequireSuccess(data);
        AssertRegionRows(data.Values, 3);
        Assert.Equal("Total Sales", data.Values[0][1]);
        Assert.Equal("Average Sales", data.Values[0][2]);
        Assert.Equal(0.25d, Convert.ToDouble(data.Values[1][1], CultureInfo.InvariantCulture));
        Assert.Equal(0.75d, Convert.ToDouble(data.Values[2][1], CultureInfo.InvariantCulture));
        Assert.Equal(100d, Convert.ToDouble(data.Values[1][2], CultureInfo.InvariantCulture));
        Assert.Equal(300d, Convert.ToDouble(data.Values[2][2], CultureInfo.InvariantCulture));
        Assert.Equal(200d, Convert.ToDouble(data.Values[^1][2], CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData("RunningTotal", null, null, null, "Total Sales")]
    [InlineData("DifferenceFrom", "Region", null, null, "Total Sales")]
    [InlineData("DifferenceFrom", "Region", "Named", null, "Total Sales")]
    [InlineData("DifferenceFrom", "Region", "Named", "Missing", "Total Sales")]
    [InlineData("DifferenceFrom", "Region", "Previous", "North", "Total Sales")]
    [InlineData("RunningTotal", "Sales", null, null, "Total Sales")]
    [InlineData("Normal", "Region", null, null, "Total Sales")]
    [InlineData("PercentOfTotal", null, null, null, "Sales")]
    [InlineData("Invented", null, null, null, "Total Sales")]
    [InlineData("DifferenceFrom", "Region", "Named", "(next)", "Total Sales")]
    [InlineData("PercentOfParentColumn", null, null, null, "Total Sales")]
    public async Task InvalidSettings_FailBeforeChangingCalculation(
        string calculation, string? baseFieldName, string? baseItemKind, string? baseItemName, string fieldName)
    {
        var name = CreatePivot();
        var failure = await _fixture.SendForFailureAsync("pivottablefield.set-field-calculation", new
        {
            pivotTableName = name,
            fieldName,
            calculation,
            baseFieldName,
            baseItemKind,
            baseItemName
        });
        Assert.False(failure.Success);
        Assert.False(string.IsNullOrWhiteSpace(failure.ErrorMessage));
        AssertValues(name, 100d, 300d);
        var listed = _fixture.Send("pivottablefield.list-fields", new { pivotTableName = name });
        using var result = JsonDocument.Parse(listed.Result!);
        Assert.Equal("Normal", Assert.Single(result.RootElement.GetProperty("valueFields").EnumerateArray())
            .GetProperty("calculation").GetString());
    }

    [Fact]
    public void ListFields_WithoutValues_ReturnsEmptyValueInstances()
    {
        var name = CreatePivot(withValue: false);
        var listed = _fixture.Send("pivottablefield.list-fields", new { pivotTableName = name });
        using var result = JsonDocument.Parse(listed.Result!);
        Assert.Equal(0, result.RootElement.GetProperty("valueFields").GetArrayLength());
    }

    private void AssertValues(string name, double north, double south)
    {
        var data = _pivot.GetData(_fixture.BatchToken, name);
        Assert.True(data.Success, data.ErrorMessage);
        AssertRegionRows(data.Values);
        Assert.NotNull(data.Values[1][1]);
        Assert.NotNull(data.Values[2][1]);
        Assert.Equal(north, Convert.ToDouble(data.Values[1][1], CultureInfo.InvariantCulture), precision: 8);
        Assert.Equal(south, Convert.ToDouble(data.Values[2][1], CultureInfo.InvariantCulture), precision: 8);
    }

    private static void AssertRegionRows(List<List<object?>> values, int columns = 2)
    {
        Assert.Equal(4, values.Count);
        Assert.All(values, row => Assert.Equal(columns, row.Count));
        Assert.Equal("North", values[1][0]);
        Assert.Equal("South", values[2][0]);
    }

    private string CreatePivot(bool withValue = true)
    {
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        var name = $"Calc_{Guid.NewGuid():N}";
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheet, "A1:B3",
            [["Region", "Sales"], ["North", 100], ["South", 300]]).Success);
        var created = _pivot.CreateFromRange(_fixture.BatchToken, sheet, "A1:B3", sheet, "D1", name);
        Assert.True(created.Success, created.ErrorMessage);
        var row = _pivot.AddRowField(_fixture.BatchToken, name, "Region");
        Assert.True(row.Success, row.ErrorMessage);
        if (withValue)
        {
            var value = _pivot.AddValueField(_fixture.BatchToken, name, "Sales", AggregationFunction.Sum, "Total Sales");
            Assert.True(value.Success, value.ErrorMessage);
        }
        return name;
    }

    private JsonDocument Set(string name, string calculation, string? baseFieldName = null,
        string? baseItemKind = null, string? baseItemName = null)
    {
        var response = _fixture.Send("pivottablefield.set-field-calculation", new
        {
            pivotTableName = name,
            fieldName = "Total Sales",
            calculation,
            baseFieldName,
            baseItemKind,
            baseItemName
        });
        var result = JsonDocument.Parse(response.Result!);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
        return result;
    }
}
