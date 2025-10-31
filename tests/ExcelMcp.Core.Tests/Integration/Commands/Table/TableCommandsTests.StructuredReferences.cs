using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

/// <summary>
/// Tests for Table structured reference operations
/// </summary>
public partial class TableCommandsTests
{
    [Fact]
    public async Task GetStructuredReference_DataRegion_ReturnsCorrectReference()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(GetStructuredReference_DataRegion_ReturnsCorrectReference));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "SalesTable", TableRegion.Data, null);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("SalesTable", result.TableName);
        Assert.Equal(TableRegion.Data, result.Region);
        Assert.Equal("SalesTable[#Data]", result.StructuredReference);
        Assert.NotNull(result.RangeAddress);
        Assert.Contains("$A$2", result.RangeAddress); // Excel returns absolute references
        Assert.Equal(4, result.RowCount); // 4 data rows
        Assert.Equal(4, result.ColumnCount); // 4 columns
    }

    [Fact]
    public async Task GetStructuredReference_AllRegion_IncludesHeaders()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(GetStructuredReference_AllRegion_IncludesHeaders));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "SalesTable", TableRegion.All, null);

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("SalesTable[#All]", result.StructuredReference);
        Assert.Contains("$A$1", result.RangeAddress); // Excel returns absolute references
        Assert.Equal(5, result.RowCount); // Headers + 4 data rows
    }

    [Fact]
    public async Task GetStructuredReference_DataRegionWithColumn_ReturnsColumnReference()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(GetStructuredReference_DataRegionWithColumn_ReturnsColumnReference));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "SalesTable", TableRegion.Data, "Amount");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("SalesTable[[Amount]]", result.StructuredReference);
        Assert.Equal("Amount", result.ColumnName);
        Assert.Equal(4, result.RowCount); // 4 data rows
        Assert.Equal(1, result.ColumnCount); // Single column
    }

    [Fact]
    public async Task GetStructuredReference_InvalidTable_ReturnsError()
    {
        // Arrange
        var testFile = await CreateTestFileWithTableAsync(nameof(GetStructuredReference_InvalidTable_ReturnsError));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _tableCommands.GetStructuredReferenceAsync(batch, "NonExistentTable", TableRegion.Data, null);

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
        Assert.Contains("not found", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }
}
