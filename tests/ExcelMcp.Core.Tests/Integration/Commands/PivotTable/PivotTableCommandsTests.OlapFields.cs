using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Tests for OLAP/Data Model PivotTable field operations (Strategy Pattern: OlapPivotTableFieldStrategy).
/// Verifies that all field manipulation methods work correctly with Data Model PivotTables.
/// Uses CubeFields API via GetFieldForManipulation() helper.
/// Organized as partial class for consistency with Strategy Pattern architecture.
/// </summary>
public partial class PivotTableCommandsTests
{
    /// <summary>
    /// OLAP-specific tests use fixture to provide Data Model PivotTable.
    /// All OLAP tests marked with [Trait("Category", "OLAP")] for strategy classification.
    /// </summary>

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public async Task AddRowField_OlapPivot_AddsFieldToRows()
    {
        // Arrange - Create OLAP test file with data model
        var olapTestFile = await CreateOlapTestFileAsync(nameof(AddRowField_OlapPivot_AddsFieldToRows));
        await using var batch = await ExcelSession.BeginBatchAsync(olapTestFile);

        // Act - Remove existing Region field first, then add Quarter
        await _pivotCommands.RemoveFieldAsync(batch, "DataModelPivot", "Region");
        var result = await _pivotCommands.AddRowFieldAsync(batch, "DataModelPivot", "Quarter", null);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Quarter", result.FieldName);
        Assert.Equal(PivotFieldArea.Row, result.Area);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public async Task AddColumnField_OlapPivot_AddsFieldToColumns()
    {
        // Arrange - Create OLAP test file with data model
        var olapTestFile = await CreateOlapTestFileAsync(nameof(AddColumnField_OlapPivot_AddsFieldToColumns));
        await using var batch = await ExcelSession.BeginBatchAsync(olapTestFile);

        // Act - Remove existing Region field first to make room for Quarter
        await _pivotCommands.RemoveFieldAsync(batch, "DataModelPivot", "Region");
        var result = await _pivotCommands.AddColumnFieldAsync(batch, "DataModelPivot", "Quarter", null);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Quarter", result.FieldName);
        Assert.Equal(PivotFieldArea.Column, result.Area);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    [Trait("Category", "OLAP")]
    public async Task SortField_OlapPivot_SortsFieldSuccessfully()
    {
        // Arrange - Create OLAP test file with data model
        var olapTestFile = await CreateOlapTestFileAsync(nameof(SortField_OlapPivot_SortsFieldSuccessfully));
        await using var batch = await ExcelSession.BeginBatchAsync(olapTestFile);

        // Act - Region row field exists in fixture
        var result = await _pivotCommands.SortFieldAsync(
            batch,
            "DataModelPivot",
            "Region",
            SortDirection.Descending);

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal("Region", result.FieldName);
    }

    /// <summary>
    /// Helper to create OLAP test file with Data Model PivotTable.
    /// Uses PivotTableRealisticFixture internally.
    /// </summary>
    private async Task<string> CreateOlapTestFileAsync(string _)
    {
        // For OLAP tests, we use the realistic fixture which has a Data Model PivotTable
        var fixture = new PivotTableRealisticFixture();
        await fixture.InitializeAsync();
        _createdFixtures.Add(fixture);
        return fixture.TestFilePath;
    }

    private readonly List<PivotTableRealisticFixture> _createdFixtures = new();
}
