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
    /// <inheritdoc/>
    [Fact]
    public void CreateFromRange_PopulatedRangeWithHeaders_CreatesCorrectPivotStructure()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(CreateFromRange_PopulatedRangeWithHeaders_CreatesCorrectPivotStructure));

        // Act
        using var batch = ExcelSession.BeginBatch(testFile);
        var result = _pivotCommands.CreateFromRange(
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
    /// <inheritdoc/>

    [Fact]
    public void CreateFromTable_WithValidTable_CreatesCorrectPivotStructure()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(CreateFromTable_WithValidTable_CreatesCorrectPivotStructure));

        // Act - Use single batch for table creation and pivot creation
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create table first
        var tableCommands = new TableCommands();
        tableCommands.Create(batch, "SalesData", "SalesTable", "A1:D6", true, TableStylePresets.Medium2);  // Create throws on error

        // Create pivot from table
        var result = _pivotCommands.CreateFromTable(
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
    /// <inheritdoc/>

    [Fact]
    public void CreateFromDataModel_NoDataModel_ReturnsError()
    {
        // Arrange - Use regular file without Data Model
        var testFile = CreateTestFileWithData(nameof(CreateFromDataModel_NoDataModel_ReturnsError));

        // Act & Assert - expects exception when Data Model is empty
        using var batch = ExcelSession.BeginBatch(testFile);
        var ex = Assert.Throws<InvalidOperationException>(() => _pivotCommands.CreateFromDataModel(
            batch,
            "AnyTable",
            "SalesData",
            "F1",
            "FailedPivot"));
        Assert.Contains("Data Model does not contain any tables", ex.Message);
    }
    /// <inheritdoc/>

    [Fact]
    public void AddRowField_WithValidField_AddsFieldToRows()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(AddRowField_WithValidField_AddsFieldToRows));

        // Act - Use single batch for create and add field
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create pivot
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Add row field
        var result = _pivotCommands.AddRowField(batch, "TestPivot", "Region");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
    }
    /// <inheritdoc/>

    [Fact]
    public void ListFields_AfterCreate_ReturnsAvailableFields()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(ListFields_AfterCreate_ReturnsAvailableFields));

        // Act - Use single batch for create and list fields
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create pivot
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // List fields
        var result = _pivotCommands.ListFields(batch, "TestPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Fields);
        Assert.True(result.Fields.Count >= 4); // Region, Product, Sales, Date
    }
}
