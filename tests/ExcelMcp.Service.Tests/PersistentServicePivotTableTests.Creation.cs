using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for PivotTable creation operations
/// </summary>
public sealed partial class PersistentServicePivotTableTests
{
    /// <inheritdoc/>
    [Fact]
    public void CreateFromRange_PopulatedRangeWithHeaders_CreatesCorrectPivotStructure()
    {
        // Arrange

        // Act
        var batch = _fixture.BatchToken;
        var result = _pivotCommands.CreateFromRange(
            batch,
            _salesSheetName, "A1:D6",
            _salesSheetName, "F1",
            "TestPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("TestPivot", result.PivotTableName);
        Assert.Equal(_salesSheetName, result.SheetName);
        Assert.Equal(4, result.AvailableFields.Count);
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateFromTable_WithValidTable_CreatesCorrectPivotStructure()
    {
        // Arrange

        // Act - Use single batch for table creation and pivot creation
        var batch = _fixture.BatchToken;

        // Create table first
        _tableCommands.Create(batch, _salesSheetName, "SalesTable", "A1:D6", true, TableStylePresets.Medium2);  // Create throws on error

        // Create pivot from table
        var result = _pivotCommands.CreateFromTable(
            batch,
            "SalesTable",
            _salesSheetName, "F1",
            "TablePivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.Equal("TablePivot", result.PivotTableName);
        Assert.Equal(_salesSheetName, result.SheetName);
        Assert.Equal(4, result.AvailableFields.Count);
    }
    /// <inheritdoc/>

    [Fact]
    public void CreateFromDataModel_NoDataModel_ReturnsError()
    {
        // Arrange - Use regular file without Data Model

        // Act & Assert - expects exception when Data Model is empty
        var batch = _fixture.BatchToken;
        var ex = Assert.Throws<InvalidOperationException>(() => _pivotCommands.CreateFromDataModel(
            batch,
            "AnyTable",
            _salesSheetName,
            "F1",
            "FailedPivot"));
        Assert.Contains("Data Model does not contain any tables", ex.Message);
    }
    /// <inheritdoc/>

    [Fact]
    public void AddRowField_WithValidField_AddsFieldToRows()
    {
        // Arrange

        // Act - Use single batch for create and add field
        var batch = _fixture.BatchToken;

        // Create pivot
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
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

        // Act - Use single batch for create and list fields
        var batch = _fixture.BatchToken;

        // Create pivot
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F1", "TestPivot");
        Assert.True(createResult.Success);

        // List fields
        var result = _pivotCommands.ListFields(batch, "TestPivot");

        // Assert
        Assert.True(result.Success, $"Expected success but got error: {result.ErrorMessage}");
        Assert.NotNull(result.Fields);
        Assert.True(result.Fields.Count >= 4); // Region, Product, Sales, Date
    }
}



