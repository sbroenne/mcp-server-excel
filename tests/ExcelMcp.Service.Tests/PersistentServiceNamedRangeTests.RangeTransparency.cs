using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceNamedRangeTests
{
    [Fact]
    public void GetValues_WithNamedRange_ResolvesProperly()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var namedRange = CreateUniqueNamedRangeName();
        _parameterCommands.Create(
            batch,
            namedRange,
            $"{sheetName}!$A$1:$B$2");
        _fixture.RegisterNamedRangeForCleanup(namedRange);
        _commands.SetValues(
            batch,
            sheetName,
            "A1:B2",
            [[1, 2], [3, 4]]);

        var result = _commands.GetValues(batch, string.Empty, namedRange);

        Assert.True(result.Success);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(
            1.0,
            Convert.ToDouble(
                result.Values[0][0],
                System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            4.0,
            Convert.ToDouble(
                result.Values[1][1],
                System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void SetValues_WithNamedRange_WritesProperly()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var namedRange = CreateUniqueNamedRangeName();
        _parameterCommands.Create(
            batch,
            namedRange,
            $"{sheetName}!$A$1:$C$2");
        _fixture.RegisterNamedRangeForCleanup(namedRange);

        var result = _commands.SetValues(
            batch,
            string.Empty,
            namedRange,
            [["Product", "Qty", "Price"], ["Widget", 10, 29.99]]);

        Assert.True(result.Success);
        var readResult = _commands.GetValues(batch, sheetName, "A1:C2");
        Assert.Equal("Product", readResult.Values[0][0]);
        Assert.Equal(
            29.99,
            Convert.ToDouble(
                readResult.Values[1][2],
                System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void GetFormulas_WithNamedRange_ReturnsFormulas()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var namedRange = CreateUniqueNamedRangeName();
        _parameterCommands.Create(
            batch,
            namedRange,
            $"{sheetName}!$A$1:$B$2");
        _fixture.RegisterNamedRangeForCleanup(namedRange);
        _commands.SetValues(batch, sheetName, "A1", [[10]]);
        _commands.SetFormulas(batch, sheetName, "B1", [["=A1*2"]]);

        var result = _commands.GetFormulas(batch, string.Empty, namedRange);

        Assert.True(result.Success);
        Assert.Empty(result.Formulas[0][0]);
        Assert.Equal("=A1*2", result.Formulas[0][1]);
        Assert.Equal(
            20.0,
            Convert.ToDouble(
                result.Values[0][1],
                System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void ClearContents_WithNamedRange_ClearsData()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var namedRange = CreateUniqueNamedRangeName();
        _parameterCommands.Create(
            batch,
            namedRange,
            $"{sheetName}!$A$1:$B$2");
        _fixture.RegisterNamedRangeForCleanup(namedRange);
        _commands.SetValues(
            batch,
            string.Empty,
            namedRange,
            [[1, 2], [3, 4]]);

        var result = _commands.ClearContents(
            batch,
            string.Empty,
            namedRange);

        Assert.True(result.Success);
        var readResult = _commands.GetValues(batch, sheetName, "A1:B2");
        Assert.All(
            readResult.Values,
            row => Assert.All(row, cell => Assert.Null(cell)));
    }
}
