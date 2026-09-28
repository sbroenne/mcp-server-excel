using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed partial class PersistentServicePowerQueryEvaluateTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPowerQueryCommands _queries =
        ServiceCommandProxy.Create<IPowerQueryCommands>(fixture);

    [Fact]
    public void Evaluate_SimpleTable_ReturnsData()
    {
        const string mCode = """
            let
                Source = #table(
                    {"Name", "Value"},
                    {{"Test1", 100}, {"Test2", 200}}
                )
            in
                Source
            """;

        var result = _queries.Evaluate(_fixture.BatchToken, mCode);

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(2, result.RowCount);
        Assert.Contains("Name", result.Columns);
        Assert.Contains("Value", result.Columns);
        Assert.Equal(2, result.Rows.Count);
    }

    [Fact]
    public void Evaluate_SingleColumn_ReturnsData()
    {
        const string mCode = """
            let
                Source = #table({"SingleCol"}, {{1}, {2}, {3}})
            in
                Source
            """;

        var result = _queries.Evaluate(_fixture.BatchToken, mCode);

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.Equal(1, result.ColumnCount);
        Assert.Equal(3, result.RowCount);
        Assert.Equal("SingleCol", result.Columns[0]);
    }

    [Fact]
    public void Evaluate_VariousDataTypes_ReturnsCorrectValues()
    {
        const string mCode = """
            let
                Source = #table(
                    {"Text", "Number", "Boolean"},
                    {{"Hello", 42, true}, {"World", 3.14, false}}
                )
            in
                Source
            """;

        var result = _queries.Evaluate(_fixture.BatchToken, mCode);

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(2, result.RowCount);
        var firstRow = result.Rows[0];
        Assert.Equal("Hello", firstRow[0]?.ToString());
        Assert.Equal(
            42.0,
            Convert.ToDouble(
                firstRow[1],
                System.Globalization.CultureInfo.InvariantCulture));
        Assert.True(
            Convert.ToBoolean(
                firstRow[2],
                System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void Evaluate_EmptyMCode_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() =>
            _queries.Evaluate(_fixture.BatchToken, ""));
    }

    [Fact]
    public void Evaluate_WithTransformations_ReturnsTransformedData()
    {
        const string mCode = """
            let
                Source = #table({"Value"}, {{1}, {2}, {3}, {4}, {5}}),
                Filtered = Table.SelectRows(Source, each [Value] > 2),
                Added = Table.AddColumn(Filtered, "Doubled", each [Value] * 2)
            in
                Added
            """;

        var result = _queries.Evaluate(_fixture.BatchToken, mCode);

        Assert.True(result.Success, $"Evaluate failed: {result.ErrorMessage}");
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(3, result.RowCount);
        Assert.Contains("Doubled", result.Columns);
    }
}
