using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for OLAP/Data Model PivotTable field operations.
/// Verifies that all field manipulation methods work correctly with Data Model PivotTables.
/// Uses CubeFields API via GetFieldForManipulation() helper.
/// Uses PivotTableRealisticFixture which provides a workbook with Data Model PivotTable.
/// </summary>
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
    public async Task AddRowField_OlapPivot_AddsFieldToRows()
    {
        // Arrange - Remove existing Region field first
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Region");

        // Act
        var result = await _commands.AddRowFieldAsync(batch, "DataModelPivot", "Quarter", null);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Quarter", result.FieldName);
        Assert.Equal(PivotFieldArea.Row, result.Area);
    }

    [Fact]
    public async Task AddColumnField_OlapPivot_AddsFieldToColumns()
    {
        // Arrange - Remove existing Region field first to make room for Quarter
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Region");

        // Act - Add Quarter field to columns
        var result = await _commands.AddColumnFieldAsync(batch, "DataModelPivot", "Quarter", null);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Quarter", result.FieldName);
        Assert.Equal(PivotFieldArea.Column, result.Area);
    }

    [Fact]
    public async Task AddValueField_OlapPivot_AddsFieldToValues()
    {
        // Arrange
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);

        // Act - Add Units field with Sum aggregation
        var result = await _commands.AddValueFieldAsync(
            batch,
            "DataModelPivot",
            "Units",
            AggregationFunction.Sum,
            "Total Units");

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Units", result.FieldName);
        Assert.Equal("Total Units", result.CustomName);
        Assert.Equal(PivotFieldArea.Value, result.Area);
        Assert.Equal(AggregationFunction.Sum, result.Function);
    }

    [Fact]
    public async Task AddFilterField_OlapPivot_AddsFieldToFilters()
    {
        // Arrange - Remove existing Region row field first
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Region");

        // Act
        var result = await _commands.AddFilterFieldAsync(batch, "DataModelPivot", "Region");

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(PivotFieldArea.Filter, result.Area);
    }

    [Fact]
    public async Task RemoveField_OlapPivot_RemovesFieldSuccessfully()
    {
        // Arrange - Add Quarter field first
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Region");
        var addResult = await _commands.AddRowFieldAsync(batch, "DataModelPivot", "Quarter", null);
        Assert.True(addResult.Success);

        // Act
        var result = await _commands.RemoveFieldAsync(batch, "DataModelPivot", "Quarter");

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Quarter", result.FieldName);
    }

    [Fact]
    public async Task SetFieldName_OlapPivot_RenamesFieldSuccessfully()
    {
        // Arrange - Add value field first
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var addResult = await _commands.AddValueFieldAsync(
            batch,
            "DataModelPivot",
            "Units",
            AggregationFunction.Sum,
            "Original Name");
        Assert.True(addResult.Success);

        // Act
        var result = await _commands.SetFieldNameAsync(batch, "DataModelPivot", "Units", "Renamed Units");

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Units", result.FieldName);
        Assert.Equal("Renamed Units", result.CustomName);
    }

    [Fact]
    public async Task SetFieldFunction_OlapPivot_ChangesAggregation()
    {
        // Arrange - Add value field with Sum
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var addResult = await _commands.AddValueFieldAsync(
            batch,
            "DataModelPivot",
            "Units",
            AggregationFunction.Sum,
            "Units Total");
        Assert.True(addResult.Success);

        // Act - Change to Average
        var result = await _commands.SetFieldFunctionAsync(
            batch,
            "DataModelPivot",
            "Units",
            AggregationFunction.Average);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Units", result.FieldName);
        Assert.Equal(AggregationFunction.Average, result.Function);
    }

    [Fact]
    public async Task SetFieldFormat_OlapPivot_SetsNumberFormat()
    {
        // Arrange - Add value field first
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var addResult = await _commands.AddValueFieldAsync(
            batch,
            "DataModelPivot",
            "Units",
            AggregationFunction.Sum,
            "Units Amount");
        Assert.True(addResult.Success);

        // Act
        var result = await _commands.SetFieldFormatAsync(
            batch,
            "DataModelPivot",
            "Units",
            "#,##0");

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Units", result.FieldName);
        Assert.Equal("#,##0", result.NumberFormat);
    }

    [Fact]
    public async Task SortField_OlapPivot_SortsFieldSuccessfully()
    {
        // Arrange - Region row field exists in fixture
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);

        // Act
        var result = await _commands.SortFieldAsync(
            batch,
            "DataModelPivot",
            "Region",
            SortDirection.Descending);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
    }

    [Fact]
    public async Task SetFieldFilter_OlapPivot_FiltersFieldSuccessfully()
    {
        // Arrange - Region row field exists in fixture
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);

        // Act
        var result = await _commands.SetFieldFilterAsync(
            batch,
            "DataModelPivot",
            "Region",
            new List<string> { "North", "South" });

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
        Assert.Equal(2, result.SelectedItems.Count);
        Assert.Contains("North", result.SelectedItems);
        Assert.Contains("South", result.SelectedItems);
    }
}
