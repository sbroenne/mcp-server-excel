using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Collection definition to force OLAP field tests to run sequentially (not in parallel).
/// Required because tests modify the same Data Model PivotTable and would interfere with each other.
/// </summary>
[CollectionDefinition("OlapFieldTests", DisableParallelization = true)]
public class OlapFieldTestsDefinition
{
}

/// <summary>
/// Tests for OLAP/Data Model PivotTable field operations.
/// Verifies that all field manipulation methods work correctly with Data Model PivotTables.
/// Uses CubeFields API via GetFieldForManipulation() helper.
/// </summary>
[Collection("OlapFieldTests")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "true")]
public class PivotTableOlapFieldTests : IClassFixture<PivotTableRealisticFixture>
{
    private readonly PivotTableCommands _commands;
    private readonly string _testFile;

    public PivotTableOlapFieldTests(PivotTableRealisticFixture fixture)
    {
        _commands = new PivotTableCommands();
        _testFile = fixture.TestFilePath;
    }

    [Fact]
    public async Task AddRowField_DataModelPivotTable_AddsFieldSuccessfully()
    {
        // Arrange - Remove existing Region field first to make room for Quarter
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Region");

        // Act - Add Quarter field to rows
        var result = await _commands.AddRowFieldAsync(batch, "DataModelPivot", "Quarter", null);

        // Assert
        Assert.True(result.Success, $"AddRowField failed: {result.ErrorMessage}");
        Assert.Equal("Quarter", result.FieldName);
        Assert.Equal(PivotFieldArea.Row, result.Area);

        // Verify field was actually added by listing fields
        var listResult = await _commands.ListFieldsAsync(batch, "DataModelPivot");
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Fields, f => f.Name.Contains("Quarter") && f.Area == PivotFieldArea.Row);
    }

    [Fact]
    public async Task AddColumnField_DataModelPivotTable_AddsFieldSuccessfully()
    {
        // Arrange - Remove existing Region field first to make room for Quarter
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Region");

        // Act - Add Quarter field to columns (Quarter is available in RegionalSalesTable)
        var result = await _commands.AddColumnFieldAsync(batch, "DataModelPivot", "Quarter", null);

        // Assert
        Assert.True(result.Success, $"AddColumnField failed: {result.ErrorMessage}");
        Assert.Equal("Quarter", result.FieldName);
        Assert.Equal(PivotFieldArea.Column, result.Area);

        // Verify field was actually added
        var listResult = await _commands.ListFieldsAsync(batch, "DataModelPivot");
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Fields, f => f.Name.Contains("Quarter") && f.Area == PivotFieldArea.Column);
    }

    [Fact]
    public async Task AddValueField_DataModelPivotTable_AddsFieldSuccessfully()
    {
        // Arrange
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);

        // Act - Add Units field to values with Sum aggregation (Units is available in RegionalSalesTable)
        var result = await _commands.AddValueFieldAsync(
            batch,
            "DataModelPivot",
            "Units",
            AggregationFunction.Sum,
            "Total Units");

        // Assert
        Assert.True(result.Success, $"AddValueField failed: {result.ErrorMessage}");
        Assert.Equal("Units", result.FieldName);
        Assert.Equal("Total Units", result.CustomName);
        Assert.Equal(PivotFieldArea.Value, result.Area);
        Assert.Equal(AggregationFunction.Sum, result.Function);

        // Verify field was actually added
        var listResult = await _commands.ListFieldsAsync(batch, "DataModelPivot");
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Fields, f => f.CustomName == "Total Units" && f.Area == PivotFieldArea.Value);
    }

    [Fact]
    public async Task AddFilterField_DataModelPivotTable_AddsFieldSuccessfully()
    {
        // Arrange - Remove existing Region row field first, then add as filter
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Region");

        // Act - Add Region field to filters
        var result = await _commands.AddFilterFieldAsync(batch, "DataModelPivot", "Region");

        // Assert
        Assert.True(result.Success, $"AddFilterField failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(PivotFieldArea.Filter, result.Area);

        // Verify field was actually added
        var listResult = await _commands.ListFieldsAsync(batch, "DataModelPivot");
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Fields, f => f.Name.Contains("Region") && f.Area == PivotFieldArea.Filter);
    }

    [Fact]
    public async Task RemoveField_DataModelPivotTable_RemovesFieldSuccessfully()
    {
        // Arrange - Add a field first (Quarter is not yet in the PivotTable)
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Region"); // Clear existing
        var addResult = await _commands.AddRowFieldAsync(batch, "DataModelPivot", "Quarter", null);
        Assert.True(addResult.Success);

        // Act - Remove the field
        var result = await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Quarter");

        // Assert
        Assert.True(result.Success, $"RemoveField failed: {result.ErrorMessage}");
        Assert.Equal("Quarter", result.FieldName);

        // Verify field was actually removed
        var listResult = await _commands.ListFieldsAsync(batch, "DataModelPivot");
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Fields, f => f.Name.Contains("Quarter") && f.Area != PivotFieldArea.Hidden);
    }

    [Fact]
    public async Task SetFieldName_DataModelPivotTable_RenamesFieldSuccessfully()
    {
        // Arrange - Add a field first
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var addResult = await _commands.AddValueFieldAsync(
            batch,
            "DataModelPivot",
            "Units",
            AggregationFunction.Sum,
            "Original Name");
        Assert.True(addResult.Success);

        // Act - Rename the field
        var result = await _commands.SetFieldNameAsync(batch, "DataModelPivot", "Units", "Renamed Units");

        // Assert
        Assert.True(result.Success, $"SetFieldName failed: {result.ErrorMessage}");
        Assert.Equal("Units", result.FieldName);
        Assert.Equal("Renamed Units", result.CustomName);

        // Verify name was actually changed
        var listResult = await _commands.ListFieldsAsync(batch, "DataModelPivot");
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Fields, f => f.CustomName == "Renamed Units");
        Assert.DoesNotContain(listResult.Fields, f => f.CustomName == "Original Name");
    }

    [Fact]
    public async Task SetFieldFunction_DataModelPivotTable_ChangesAggregationSuccessfully()
    {
        // Arrange - Add a value field first with Sum
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var addResult = await _commands.AddValueFieldAsync(
            batch,
            "DataModelPivot",
            "Units",
            AggregationFunction.Sum,
            "Units Total");
        Assert.True(addResult.Success);

        // Act - Change aggregation to Average
        var result = await _commands.SetFieldFunctionAsync(
            batch,
            "DataModelPivot",
            "Units",
            AggregationFunction.Average);

        // Assert
        Assert.True(result.Success, $"SetFieldFunction failed: {result.ErrorMessage}");
        Assert.Equal("Units", result.FieldName);
        Assert.Equal(AggregationFunction.Average, result.Function);
    }

    [Fact]
    public async Task SetFieldFormat_DataModelPivotTable_SetsNumberFormatSuccessfully()
    {
        // Arrange - Add a value field first
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var addResult = await _commands.AddValueFieldAsync(
            batch,
            "DataModelPivot",
            "Units",
            AggregationFunction.Sum,
            "Units Amount");
        Assert.True(addResult.Success);

        // Act - Set number format
        var result = await _commands.SetFieldFormatAsync(
            batch,
            "DataModelPivot",
            "Units",
            "#,##0");

        // Assert
        Assert.True(result.Success, $"SetFieldFormat failed: {result.ErrorMessage}");
        Assert.Equal("Units", result.FieldName);
        Assert.Equal("#,##0", result.NumberFormat);
    }

    [Fact]
    public async Task SortField_DataModelPivotTable_SortsFieldSuccessfully()
    {
        // Arrange - Region row field already exists in fixture, use it directly
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);

        // Act - Sort descending (Region is already a row field from fixture)
        var result = await _commands.SortFieldAsync(
            batch,
            "DataModelPivot",
            "Region",
            SortDirection.Descending);

        // Assert
        Assert.True(result.Success, $"SortField failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
    }

    [Fact]
    public async Task SetFieldFilter_DataModelPivotTable_FiltersFieldSuccessfully()
    {
        // Arrange - Region row field already exists in fixture, use it directly
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);

        // Act - Filter to show only North and South (Region is already a row field from fixture)
        var result = await _commands.SetFieldFilterAsync(
            batch,
            "DataModelPivot",
            "Region",
            new List<string> { "North", "South" });

        // Assert
        Assert.True(result.Success, $"SetFieldFilter failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(2, result.SelectedItems.Count);
        Assert.Contains("North", result.SelectedItems);
        Assert.Contains("South", result.SelectedItems);
    }

    [Fact]
    public async Task MultipleFieldOperations_DataModelPivotTable_WorkSequentially()
    {
        // This test verifies that multiple field operations work in sequence
        // Tests the complete workflow: Add → Configure → Rename → Remove

        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);

        // Clear existing fields to start fresh
        await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Region");

        // 1. Add row field
        var addRow = await _commands.AddRowFieldAsync(batch, "DataModelPivot", "Quarter", null);
        Assert.True(addRow.Success, $"Add row failed: {addRow.ErrorMessage}");

        // 2. Add value field
        var addValue = await _commands.AddValueFieldAsync(
            batch,
            "DataModelPivot",
            "Units",
            AggregationFunction.Sum,
            "Q Units");
        Assert.True(addValue.Success, $"Add value failed: {addValue.ErrorMessage}");

        // 3. Rename value field
        var rename = await _commands.SetFieldNameAsync(batch, "DataModelPivot", "Units", "Quarterly Units");
        Assert.True(rename.Success, $"Rename failed: {rename.ErrorMessage}");

        // 4. Set format
        var format = await _commands.SetFieldFormatAsync(batch, "DataModelPivot", "Units", "#,##0");
        Assert.True(format.Success, $"Format failed: {format.ErrorMessage}");

        // 5. Sort row field
        var sort = await _commands.SortFieldAsync(batch, "DataModelPivot", "Quarter", SortDirection.Descending);
        Assert.True(sort.Success, $"Sort failed: {sort.ErrorMessage}");

        // 6. Remove row field
        var remove = await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Quarter");
        Assert.True(remove.Success, $"Remove failed: {remove.ErrorMessage}");

        // Verify final state
        var listResult = await _commands.ListFieldsAsync(batch, "DataModelPivot");
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Fields, f => f.Name.Contains("Quarter") && f.Area != PivotFieldArea.Hidden);
        Assert.Contains(listResult.Fields, f => f.CustomName == "Quarterly Units");
    }
}
