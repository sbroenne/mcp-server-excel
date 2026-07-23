using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for named range transparency - verifying that RangeCommands works seamlessly with named ranges
/// </summary>
public partial class RangeCommandsTests
{
    // === NAMED RANGE TRANSPARENCY TESTS ===

    [Fact]
    public void GetValues_WithNamedRange_ResolvesProperly()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Create a named range pointing to A1:B2
        var paramCommands = new NamedRangeCommands();
        paramCommands.Create(batch, "TestData", $"{sheetName}!$A$1:$B$2");

        // Set data in the range
        _commands.SetValues(batch, sheetName, "A1:B2",
        [
            [1, 2],
            [3, 4]
        ]);

        // Act - Read using named range (empty sheetName)
        var result = _commands.GetValues(batch, "", "TestData");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(
            1.0,
            Convert.ToDouble(result.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            4.0,
            Convert.ToDouble(result.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));

        // Cleanup named range
        paramCommands.Delete(batch, "TestData");
    }

    [Fact]
    public void SetValues_WithNamedRange_WritesProperly()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Create a named range
        var paramCommands = new NamedRangeCommands();
        paramCommands.Create(batch, "SalesData", $"{sheetName}!$A$1:$C$2");

        // Act - Write using named range
        var result = _commands.SetValues(batch, "", "SalesData",
        [
            ["Product", "Qty", "Price"],
            ["Widget", 10, 29.99]
        ]);

        // Assert
        Assert.True(result.Success);

        // Verify by reading with regular range address
        var readResult = _commands.GetValues(batch, sheetName, "A1:C2");
        Assert.Equal("Product", readResult.Values[0][0]);
        Assert.Equal(
            29.99,
            Convert.ToDouble(readResult.Values[1][2], System.Globalization.CultureInfo.InvariantCulture));

        // Cleanup named range
        paramCommands.Delete(batch, "SalesData");
    }

    [Fact]
    public void GetFormulas_WithNamedRange_ReturnsFormulas()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Create named range and set data + formula
        var paramCommands = new NamedRangeCommands();
        paramCommands.Create(batch, "CalcRange", $"{sheetName}!$A$1:$B$2");

        _commands.SetValues(batch, sheetName, "A1", [[10]]);
        _commands.SetFormulas(batch, sheetName, "B1", [["=A1*2"]]);

        // Act - Read formulas using named range
        var result = _commands.GetFormulas(batch, "", "CalcRange");

        // Assert
        Assert.True(result.Success);
        Assert.Empty(result.Formulas[0][0]); // A1 has no formula
        Assert.Equal("=A1*2", result.Formulas[0][1]);
        Assert.Equal(
            20.0,
            Convert.ToDouble(result.Values[0][1], System.Globalization.CultureInfo.InvariantCulture));

        // Cleanup named range
        paramCommands.Delete(batch, "CalcRange");
    }

    [Fact]
    public void ClearContents_WithNamedRange_ClearsData()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Create named range and populate
        var paramCommands = new NamedRangeCommands();
        paramCommands.Create(batch, "TempData", $"{sheetName}!$A$1:$B$2");

        _commands.SetValues(batch, "", "TempData",
        [
            [1, 2],
            [3, 4]
        ]);

        // Act - Clear using named range
        var result = _commands.ClearContents(batch, "", "TempData");

        // Assert
        Assert.True(result.Success);

        // Verify data is cleared
        var readResult = _commands.GetValues(batch, sheetName, "A1:B2");
        Assert.All(readResult.Values, row => Assert.All(row, cell => Assert.Null(cell)));

        // Cleanup named range
        paramCommands.Delete(batch, "TempData");
    }
}




