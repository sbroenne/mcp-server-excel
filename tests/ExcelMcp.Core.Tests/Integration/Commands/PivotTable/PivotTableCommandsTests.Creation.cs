using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for PivotTable creation operations
/// </summary>
public partial class PivotTableCommandsTests
{
    [Fact]
    public async Task CreateFromRange_PopulatedRangeWithHeaders_CreatesCorrectPivotStructure()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(CreateFromRange_PopulatedRangeWithHeaders_CreatesCorrectPivotStructure));

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _pivotCommands.CreateFromRangeAsync(
            batch,
            "SalesData", "A1:D6",
            "SalesData", "F1",
            "TestPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.Equal("SalesData", result.SheetName);
        Assert.Equal(4, result.AvailableFields.Count);
    }

    [Fact]
    public async Task CreateFromTable_WithValidTable_CreatesCorrectPivotStructure()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(CreateFromTable_WithValidTable_CreatesCorrectPivotStructure));

        // Act - Use single batch for table creation and pivot creation
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create table first
        var tableCommands = new TableCommands();
        var tableResult = await tableCommands.CreateAsync(batch, "SalesData", "SalesTable", "A1:D6", true, "TableStyleMedium2");
        Assert.True(tableResult.Success, $"Table creation failed: {tableResult.ErrorMessage}");

        // Create pivot from table
        var result = await _pivotCommands.CreateFromTableAsync(
            batch,
            "SalesTable",
            "SalesData", "F1",
            "TablePivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("TablePivot", result.PivotTableName);
        Assert.Equal("SalesData", result.SheetName);
        Assert.Equal(4, result.AvailableFields.Count);
    }

    [Fact]
    public async Task AddRowField_WithValidField_AddsFieldToRows()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(AddRowField_WithValidField_AddsFieldToRows));

        // Act - Use single batch for create and add field
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create pivot
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add row field
        var result = await _pivotCommands.AddRowFieldAsync(batch, "TestPivot", "Region");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
    }

    [Fact]
    public async Task ListFields_AfterCreate_ReturnsAvailableFields()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(ListFields_AfterCreate_ReturnsAvailableFields));

        // Act - Use single batch for create and list fields
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Create pivot
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // List fields
        var result = await _pivotCommands.ListFieldsAsync(batch, "TestPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Fields);
        Assert.True(result.Fields.Count >= 4); // Region, Product, Sales, Date
    }
}
