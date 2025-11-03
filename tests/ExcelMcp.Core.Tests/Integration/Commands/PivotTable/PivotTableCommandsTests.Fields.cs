using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for PivotTable field operations (AddColumn, AddValue, AddFilter, Remove, Set*)
/// Optimized: Single batch per test, no SaveAsync() unless testing persistence
/// </summary>
public partial class PivotTableCommandsTests
{
    [Fact]
    [Trait("Speed", "Medium")]
    public async Task AddColumnField_WithValidField_AddsFieldToColumns()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(AddColumnField_WithValidField_AddsFieldToColumns));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act - No save needed
        var result = await _pivotCommands.AddColumnFieldAsync(batch, "TestPivot", "Product");

        // Assert
        Assert.True(result.Success, $"AddColumnField failed: {result.ErrorMessage}");
        Assert.Equal("Product", result.FieldName);
        Assert.Equal(PivotFieldArea.Column, result.Area);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task AddValueField_WithValidField_AddsFieldToValues()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(AddValueField_WithValidField_AddsFieldToValues));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act
        var result = await _pivotCommands.AddValueFieldAsync(batch, "TestPivot", "Sales");

        // Assert
        Assert.True(result.Success, $"AddValueField failed: {result.ErrorMessage}");
        Assert.Equal("Sales", result.FieldName);
        Assert.Equal(PivotFieldArea.Value, result.Area);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task AddFilterField_WithValidField_AddsFieldToFilters()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(AddFilterField_WithValidField_AddsFieldToFilters));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);

        // Act
        var result = await _pivotCommands.AddFilterFieldAsync(batch, "TestPivot", "Region");

        // Assert
        Assert.True(result.Success, $"AddFilterField failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(PivotFieldArea.Filter, result.Area);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task RemoveField_ExistingField_RemovesFromPivot()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(RemoveField_ExistingField_RemovesFromPivot));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);
        
        // Add a field first
        var addResult = await _pivotCommands.AddRowFieldAsync(batch, "TestPivot", "Region");
        Assert.True(addResult.Success);

        // Act - Remove in same batch
        var result = await _pivotCommands.RemoveFieldAsync(batch, "TestPivot", "Region");

        // Assert
        Assert.True(result.Success, $"RemoveField failed: {result.ErrorMessage}");
        
        // Verify field removed
        var infoResult = await _pivotCommands.GetInfoAsync(batch, "TestPivot");
        Assert.True(infoResult.Success);
        var regionField = infoResult.Fields.FirstOrDefault(f => f.Name == "Region");
        Assert.NotNull(regionField);
        Assert.Equal(PivotFieldArea.Hidden, regionField.Area);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task SetFieldFunction_ValueField_ChangesAggregation()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(SetFieldFunction_ValueField_ChangesAggregation));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);
        
        // Add Sales as value field (default sum)
        var addResult = await _pivotCommands.AddValueFieldAsync(batch, "TestPivot", "Sales");
        Assert.True(addResult.Success);

        // Act - Change to Average in same batch
        var result = await _pivotCommands.SetFieldFunctionAsync(batch, "TestPivot", "Sales", AggregationFunction.Average);

        // Assert
        Assert.True(result.Success, $"SetFieldFunction failed: {result.ErrorMessage}");
        Assert.Equal("Sales", result.FieldName);
        Assert.Equal(AggregationFunction.Average, result.Function);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task SetFieldName_ExistingField_RenamesField()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(SetFieldName_ExistingField_RenamesField));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);
        
        // Add Sales as value field
        var addResult = await _pivotCommands.AddValueFieldAsync(batch, "TestPivot", "Sales");
        Assert.True(addResult.Success);

        // Act
        var result = await _pivotCommands.SetFieldNameAsync(batch, "TestPivot", "Sales", "Total Revenue");

        // Assert
        Assert.True(result.Success, $"SetFieldName failed: {result.ErrorMessage}");
        Assert.Equal("Total Revenue", result.CustomName);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task SetFieldFormat_ValueField_AppliesNumberFormat()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(SetFieldFormat_ValueField_AppliesNumberFormat));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);
        
        // Add Sales as value field
        var addResult = await _pivotCommands.AddValueFieldAsync(batch, "TestPivot", "Sales");
        Assert.True(addResult.Success);

        // Act
        var result = await _pivotCommands.SetFieldFormatAsync(batch, "TestPivot", "Sales", "$#,##0.00");

        // Assert
        Assert.True(result.Success, $"SetFieldFormat failed: {result.ErrorMessage}");
        Assert.Equal("$#,##0.00", result.NumberFormat);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task SetFieldFilter_RowField_AppliesFilter()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(SetFieldFilter_RowField_AppliesFilter));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);
        
        // Add Region as row field
        var addResult = await _pivotCommands.AddRowFieldAsync(batch, "TestPivot", "Region");
        Assert.True(addResult.Success);

        // Act
        var result = await _pivotCommands.SetFieldFilterAsync(batch, "TestPivot", "Region", new List<string> { "North" });

        // Assert
        Assert.True(result.Success, $"SetFieldFilter failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.NotEmpty(result.SelectedItems);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public async Task SortField_RowField_SortsData()
    {
        // Arrange
        var testFile = await CreateTestFileWithDataAsync(nameof(SortField_RowField_SortsData));

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var createResult = await _pivotCommands.CreateFromRangeAsync(
            batch, "SalesData", "A1:D6", "SalesData", "F1", "TestPivot");
        Assert.True(createResult.Success);
        
        // Add Region as row field
        var addResult = await _pivotCommands.AddRowFieldAsync(batch, "TestPivot", "Region");
        Assert.True(addResult.Success);

        // Act
        var result = await _pivotCommands.SortFieldAsync(batch, "TestPivot", "Region", SortDirection.Ascending);

        // Assert
        Assert.True(result.Success, $"SortField failed: {result.ErrorMessage}");
    }
}
