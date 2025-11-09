using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.PivotTable;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

/// <summary>
/// Realistic scenario tests for PivotTable operations using shared fixture.
/// Tests PivotTables from multiple sources: ranges, Excel tables, and Data Model.
/// CRITICAL: Tests the bug where List operation fails on real workbooks with defensive error handling.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "PivotTables")]
[Trait("RequiresExcel", "true")]
public class PivotTableRealisticTests : IClassFixture<PivotTableRealisticFixture>
{
    private readonly PivotTableCommands _commands;
    private readonly string _testFile;

    public PivotTableRealisticTests(PivotTableRealisticFixture fixture)
    {
        _commands = new PivotTableCommands();
        _testFile = fixture.TestFilePath;
    }

    /// <summary>
    /// Tests that List operation works with PivotTables from multiple source types.
    /// This reproduces the user's bug where List failed with 0x800A03EC on real workbooks.
    /// </summary>
    [Fact]
    public async Task List_RealisticWorkbook_ReturnsAllPivotTables()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.PivotTables);
        Assert.Equal(3, result.PivotTables.Count);

        // Verify all three PivotTable types are listed
        Assert.Contains(result.PivotTables, pt => pt.Name == "SalesByRegion");
        Assert.Contains(result.PivotTables, pt => pt.Name == "RegionalSummary");
        Assert.Contains(result.PivotTables, pt => pt.Name == "DataModelPivot");
    }

    /// <summary>
    /// Tests Get operation on range-based PivotTable
    /// </summary>
    [Fact]
    public async Task Get_RangePivotTable_ReturnsDetails()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.GetAsync(batch, "SalesByRegion");

        // Assert
        Assert.True(result.Success, $"Get failed: {result.ErrorMessage}");
        Assert.Equal("SalesByRegion", result.PivotTable.Name);
        Assert.Equal("PivotData", result.PivotTable.SheetName);
        Assert.NotNull(result.PivotTable.SourceData);
    }

    /// <summary>
    /// Tests Get operation on table-based PivotTable
    /// </summary>
    [Fact]
    public async Task Get_TablePivotTable_ReturnsDetails()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.GetAsync(batch, "RegionalSummary");

        // Assert
        Assert.True(result.Success, $"Get failed: {result.ErrorMessage}");
        Assert.Equal("RegionalSummary", result.PivotTable.Name);
        Assert.Equal("RegionalData", result.PivotTable.SheetName);
        Assert.NotNull(result.PivotTable.SourceData);
    }

    /// <summary>
    /// Tests Get operation on Data Model PivotTable.
    /// This type commonly had inaccessible properties that caused List to fail.
    /// </summary>
    [Fact]
    public async Task Get_DataModelPivotTable_ReturnsDetails()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.GetAsync(batch, "DataModelPivot");

        // Assert
        Assert.True(result.Success, $"Get failed: {result.ErrorMessage}");
        Assert.Equal("DataModelPivot", result.PivotTable.Name);
        Assert.Equal("ModelData", result.PivotTable.SheetName);

        // Data Model PivotTables may not have accessible SourceData
        // The defensive error handling should handle this gracefully
    }

    /// <summary>
    /// Tests ListFields on range-based PivotTable
    /// </summary>
    [Fact]
    public async Task ListFields_RangePivotTable_ReturnsFields()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListFieldsAsync(batch, "SalesByRegion");

        // Assert
        Assert.True(result.Success, $"ListFields failed: {result.ErrorMessage}");
        Assert.NotNull(result.Fields);
        Assert.Contains(result.Fields, f => f.Name == "Region");
        Assert.Contains(result.Fields, f => f.Name == "Revenue");
    }

    /// <summary>
    /// Tests ListFields on Data Model PivotTable
    /// </summary>
    [Fact]
    public async Task ListFields_DataModelPivotTable_ReturnsFields()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListFieldsAsync(batch, "DataModelPivot");

        // Assert
        Assert.True(result.Success, $"ListFields failed: {result.ErrorMessage}");
        Assert.NotNull(result.Fields);

        // OLAP/Data Model PivotTables use CubeFields with full hierarchy names
        Assert.Contains(result.Fields, f => f.Name.Contains("Region"));
    }

    /// <summary>
    /// Tests Refresh on multiple PivotTable types.
    /// Verifies defensive error handling works during refresh operations.
    /// </summary>
    [Fact]
    public async Task Refresh_AllPivotTableTypes_Succeeds()
    {
        // Arrange
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);

        // Act & Assert - Range-based
        var rangeResult = await _commands.RefreshAsync(batch, "SalesByRegion");
        Assert.True(rangeResult.Success, $"Range refresh failed: {rangeResult.ErrorMessage}");

        // Act & Assert - Table-based
        var tableResult = await _commands.RefreshAsync(batch, "RegionalSummary");
        Assert.True(tableResult.Success, $"Table refresh failed: {tableResult.ErrorMessage}");

        // Act & Assert - Data Model
        var modelResult = await _commands.RefreshAsync(batch, "DataModelPivot");
        Assert.True(modelResult.Success, $"Data Model refresh failed: {modelResult.ErrorMessage}");
    }

    /// <summary>
    /// Tests that defensive error handling allows partial success.
    /// Even if one PivotTable has issues, others should still be listed.
    /// </summary>
    [Fact]
    public async Task List_WithMixedAccessibility_ReturnsAccessiblePivotTables()
    {
        // This test validates the defensive error handling pattern:
        // - If a PivotTable property throws an exception, we catch it
        // - Continue processing other PivotTables
        // - Return all successfully processed PivotTables

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List should succeed even with some property access failures: {result.ErrorMessage}");
        Assert.NotNull(result.PivotTables);

        // We should get at least some PivotTables back, even if not all properties are accessible
        Assert.NotEmpty(result.PivotTables);
    }
}
