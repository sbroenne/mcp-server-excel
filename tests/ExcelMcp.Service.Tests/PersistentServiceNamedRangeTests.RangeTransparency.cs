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
        Assert.True(_parameterCommands.Create(
            batch,
            namedRange,
            $"{sheetName}!$A$1:$B$2").Success);
        _fixture.RegisterNamedRangeForCleanup(namedRange);
        Assert.True(_commands.SetValues(
            batch,
            sheetName,
            "A1:B2",
            [[1, 2], [3, 4]]).Success);

        var result = _commands.GetValues(batch, string.Empty, namedRange);

        Assert.True(result.Success);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(2, result.Values.Count);
        for (var row = 0; row < 2; row++)
        {
            Assert.Equal(2, result.Values[row].Count);
            for (var column = 0; column < 2; column++)
            {
                Assert.Equal(row * 2 + column + 1,
                    Convert.ToDouble(result.Values[row][column], System.Globalization.CultureInfo.InvariantCulture));
            }
        }
    }

    [Fact]
    public void SetValues_WithNamedRange_WritesProperly()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var namedRange = CreateUniqueNamedRangeName();
        Assert.True(_parameterCommands.Create(
            batch,
            namedRange,
            $"{sheetName}!$A$1:$C$2").Success);
        _fixture.RegisterNamedRangeForCleanup(namedRange);

        var result = _commands.SetValues(
            batch,
            string.Empty,
            namedRange,
            [["Product", "Qty", "Price"], ["Widget", 10, 29.99]]);

        Assert.True(result.Success);
        var readResult = _commands.GetValues(batch, sheetName, "A1:C2");
        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.Equal(2, readResult.RowCount);
        Assert.Equal(3, readResult.ColumnCount);
        Assert.Equal(["Product", "Qty", "Price"], readResult.Values[0]);
        Assert.Equal("Widget", readResult.Values[1][0]);
        Assert.Equal(10, Convert.ToDouble(readResult.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));
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
        Assert.True(_parameterCommands.Create(
            batch,
            namedRange,
            $"{sheetName}!$A$1:$B$2").Success);
        _fixture.RegisterNamedRangeForCleanup(namedRange);
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[10]]).Success);
        Assert.True(_commands.SetFormulas(batch, sheetName, "B1", [["=A1*2"]]).Success);

        var result = _commands.GetFormulas(batch, string.Empty, namedRange);

        Assert.True(result.Success);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Collection(result.Formulas,
            row => Assert.Equal(["", "=A1*2"], row),
            row => Assert.Equal(["", ""], row));
        Assert.Equal(10, Convert.ToDouble(result.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            20.0,
            Convert.ToDouble(
                result.Values[0][1],
                System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(2, result.Values.Count);
        Assert.Equal(2, result.Values[1].Count);
        Assert.All(result.Values[1], cell => Assert.Null(cell));
    }

    [Fact]
    public void ClearContents_WithNamedRange_ClearsData()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var namedRange = CreateUniqueNamedRangeName();
        Assert.True(_parameterCommands.Create(
            batch,
            namedRange,
            $"{sheetName}!$A$1:$B$2").Success);
        _fixture.RegisterNamedRangeForCleanup(namedRange);
        Assert.True(_commands.SetValues(
            batch,
            string.Empty,
            namedRange,
            [[1, 2], [3, 4]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "C1", [["Untouched"]]).Success);

        var result = _commands.ClearContents(
            batch,
            string.Empty,
            namedRange);

        Assert.True(result.Success);
        var readResult = _commands.GetValues(batch, sheetName, "A1:B2");
        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.Equal(2, readResult.RowCount);
        Assert.Equal(2, readResult.ColumnCount);
        Assert.Equal(2, readResult.Values.Count);
        Assert.All(readResult.Values, row => Assert.Equal(2, row.Count));
        Assert.All(
            readResult.Values,
            row => Assert.All(row, cell => Assert.Null(cell)));
        var untouched = _commands.GetValues(batch, sheetName, "C1");
        Assert.True(untouched.Success, untouched.ErrorMessage);
        Assert.Equal("Untouched", Assert.Single(Assert.Single(untouched.Values)));
    }
}
