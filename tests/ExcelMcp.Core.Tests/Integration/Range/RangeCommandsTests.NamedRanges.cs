using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Range;

/// <summary>
/// Tests for named range transparency - verifying that RangeCommands works seamlessly with named ranges
/// </summary>
public partial class RangeCommandsTests
{
    // === NAMED RANGE TRANSPARENCY TESTS ===

    [Fact]
    public async Task GetValuesAsync_WithNamedRange_ResolvesProperly()
    {
        // Arrange
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create a named range pointing to A1:B2
        var paramCommands = new ParameterCommands();
        await paramCommands.CreateAsync(batch, "TestData", "Sheet1!$A$1:$B$2");

        // Set data in the range
        await _commands.SetValuesAsync(batch, "Sheet1", "A1:B2", new List<List<object?>>
        {
            new() { 1, 2 },
            new() { 3, 4 }
        });

        // Act - Read using named range (empty sheetName)
        var result = await _commands.GetValuesAsync(batch, "", "TestData");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(1.0, Convert.ToDouble(result.Values[0][0]));
        Assert.Equal(4.0, Convert.ToDouble(result.Values[1][1]));
    }

    [Fact]
    public async Task SetValuesAsync_WithNamedRange_WritesProperly()
    {
        // Arrange
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create a named range
        var paramCommands = new ParameterCommands();
        await paramCommands.CreateAsync(batch, "SalesData", "Sheet1!$A$1:$C$2");

        // Act - Write using named range
        var result = await _commands.SetValuesAsync(batch, "", "SalesData", new List<List<object?>>
        {
            new() { "Product", "Qty", "Price" },
            new() { "Widget", 10, 29.99 }
        });
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);

        // Verify by reading with regular range address
        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "A1:C2");
        Assert.Equal("Product", readResult.Values[0][0]);
        Assert.Equal(29.99, Convert.ToDouble(readResult.Values[1][2]));
    }

    [Fact]
    public async Task GetFormulasAsync_WithNamedRange_ReturnsFormulas()
    {
        // Arrange
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create named range and set data + formula
        var paramCommands = new ParameterCommands();
        await paramCommands.CreateAsync(batch, "CalcRange", "Sheet1!$A$1:$B$2");

        await _commands.SetValuesAsync(batch, "Sheet1", "A1", new List<List<object?>> { new() { 10 } });
        await _commands.SetFormulasAsync(batch, "Sheet1", "B1", new List<List<string>> { new() { "=A1*2" } });

        // Act - Read formulas using named range
        var result = await _commands.GetFormulasAsync(batch, "", "CalcRange");

        // Assert
        Assert.True(result.Success);
        Assert.Empty(result.Formulas[0][0]); // A1 has no formula
        Assert.Equal("=A1*2", result.Formulas[0][1]);
        Assert.Equal(20.0, Convert.ToDouble(result.Values[0][1]));
    }

    [Fact]
    public async Task ClearContentsAsync_WithNamedRange_ClearsData()
    {
        // Arrange
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Create named range and populate
        var paramCommands = new ParameterCommands();
        await paramCommands.CreateAsync(batch, "TempData", "Sheet1!$A$1:$B$2");

        await _commands.SetValuesAsync(batch, "", "TempData", new List<List<object?>>
        {
            new() { 1, 2 },
            new() { 3, 4 }
        });

        // Act - Clear using named range
        var result = await _commands.ClearContentsAsync(batch, "", "TempData");
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);

        // Verify data is cleared
        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "A1:B2");
        Assert.All(readResult.Values, row => Assert.All(row, cell => Assert.Null(cell)));
    }
}
