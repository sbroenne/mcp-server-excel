using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

/// <summary>
/// Integration tests for Table Slicer operations.
/// Tests cover: create, list, set selection, delete slicers for Excel Tables.
/// Uses TableTestsFixture which creates isolated table files per test.
/// </summary>
public partial class TableCommandsTests
{
    #region Table Slicer Tests

    /// <summary>
    /// Tests creating a slicer for a Table column.
    /// </summary>
    [Fact]
    public void CreateTableSlicer_ValidColumn_CreatesSlicerSuccessfully()
    {
        // Arrange - Create a fresh test file with SalesTable
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create slicer for Region column
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch,
            tableName: "SalesTable",
            columnName: "Region",
            slicerName: "RegionSlicer",
            destinationSheet: "Sales",
            position: "F2");

        // Assert
        Assert.True(slicerResult.Success, $"CreateTableSlicer failed: {slicerResult.ErrorMessage}");
        Assert.Equal("RegionSlicer", slicerResult.Name);
        Assert.Equal("Region", slicerResult.FieldName);
        Assert.Equal("Sales", slicerResult.SheetName);
        Assert.NotNull(slicerResult.AvailableItems);
        Assert.Contains("North", slicerResult.AvailableItems);
        Assert.Contains("South", slicerResult.AvailableItems);
        Assert.Contains("East", slicerResult.AvailableItems);
        Assert.Contains("West", slicerResult.AvailableItems);
        Assert.Equal("SalesTable", slicerResult.ConnectedTable);
        Assert.Equal("Table", slicerResult.SourceType);
        Assert.NotNull(slicerResult.WorkflowHint);
    }

    /// <summary>
    /// Tests listing Table slicers in a workbook with no filter.
    /// </summary>
    [Fact]
    public void ListTableSlicers_WithSlicers_ReturnsAllSlicers()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create two slicers
        var slicer1Result = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "RegionSlicer1", "Sales", "F2");
        Assert.True(slicer1Result.Success, $"Failed to create slicer 1: {slicer1Result.ErrorMessage}");

        var slicer2Result = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Product", "ProductSlicer1", "Sales", "F10");
        Assert.True(slicer2Result.Success, $"Failed to create slicer 2: {slicer2Result.ErrorMessage}");

        // Act
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert
        Assert.True(listResult.Success, $"ListTableSlicers failed: {listResult.ErrorMessage}");
        Assert.NotNull(listResult.Slicers);
        Assert.True(listResult.Slicers.Count >= 2, $"Expected at least 2 slicers, got {listResult.Slicers.Count}");
        Assert.Contains(listResult.Slicers, s => s.Name == "RegionSlicer1");
        Assert.Contains(listResult.Slicers, s => s.Name == "ProductSlicer1");
    }

    /// <summary>
    /// Tests listing slicers filtered by Table name.
    /// </summary>
    [Fact]
    public void ListTableSlicers_FilterByTable_ReturnsConnectedSlicersOnly()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create slicer for SalesTable
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "FilterRegionSlicer", "Sales", "F2");
        Assert.True(slicerResult.Success, $"Failed to create slicer: {slicerResult.ErrorMessage}");

        // Act
        var listResult = _tableCommands.ListTableSlicers(batch, tableName: "SalesTable");

        // Assert
        Assert.True(listResult.Success, $"ListTableSlicers failed: {listResult.ErrorMessage}");
        Assert.NotNull(listResult.Slicers);
        Assert.Single(listResult.Slicers);
        Assert.Equal("FilterRegionSlicer", listResult.Slicers[0].Name);
        Assert.Equal("SalesTable", listResult.Slicers[0].ConnectedTable);
    }

    /// <summary>
    /// Tests setting Table slicer selection to specific items.
    /// </summary>
    [Fact]
    public void SetTableSlicerSelection_SpecificItems_SelectsOnlyThoseItems()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create slicer
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "SelectionSlicer", "Sales", "F2");
        Assert.True(slicerResult.Success, $"Failed to create slicer: {slicerResult.ErrorMessage}");

        // Act - Select only "North" and "South"
        var selectionResult = _tableCommands.SetTableSlicerSelection(
            batch, "SelectionSlicer", new List<string> { "North", "South" }, clearFirst: true);

        // Assert
        Assert.True(selectionResult.Success, $"SetTableSlicerSelection failed: {selectionResult.ErrorMessage}");
        Assert.NotNull(selectionResult.SelectedItems);
        Assert.Equal(2, selectionResult.SelectedItems.Count);
        Assert.Contains("North", selectionResult.SelectedItems);
        Assert.Contains("South", selectionResult.SelectedItems);
        Assert.DoesNotContain("East", selectionResult.SelectedItems);
        Assert.DoesNotContain("West", selectionResult.SelectedItems);
        Assert.NotNull(selectionResult.WorkflowHint);
    }

    /// <summary>
    /// Tests clearing Table slicer selection (selecting all items).
    /// </summary>
    [Fact]
    public void SetTableSlicerSelection_EmptyList_ClearsFilterSelectsAll()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create slicer
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "ClearFilterSlicer", "Sales", "F2");
        Assert.True(slicerResult.Success, $"Failed to create slicer: {slicerResult.ErrorMessage}");

        // First, filter to just "North"
        _tableCommands.SetTableSlicerSelection(batch, "ClearFilterSlicer", new List<string> { "North" });

        // Act - Clear filter by passing empty list
        var clearResult = _tableCommands.SetTableSlicerSelection(
            batch, "ClearFilterSlicer", new List<string>());

        // Assert
        Assert.True(clearResult.Success, $"SetTableSlicerSelection (clear) failed: {clearResult.ErrorMessage}");
        Assert.NotNull(clearResult.SelectedItems);
        Assert.True(clearResult.SelectedItems.Count >= 4, "Expected all items to be selected after clear");
        Assert.Contains("North", clearResult.SelectedItems);
        Assert.Contains("South", clearResult.SelectedItems);
        Assert.Contains("East", clearResult.SelectedItems);
        Assert.Contains("West", clearResult.SelectedItems);
        Assert.Contains("cleared", clearResult.WorkflowHint, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests deleting a Table slicer from the workbook.
    /// </summary>
    [Fact]
    public void DeleteTableSlicer_ExistingSlicer_RemovesSuccessfully()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create slicer
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "SlicerToDelete", "Sales", "F2");
        Assert.True(slicerResult.Success, $"Failed to create slicer: {slicerResult.ErrorMessage}");

        // Verify slicer exists
        var listBeforeResult = _tableCommands.ListTableSlicers(batch);
        Assert.Contains(listBeforeResult.Slicers, s => s.Name == "SlicerToDelete");

        // Act
        var deleteResult = _tableCommands.DeleteTableSlicer(batch, "SlicerToDelete");

        // Assert
        Assert.True(deleteResult.Success, $"DeleteTableSlicer failed: {deleteResult.ErrorMessage}");

        // Verify slicer is gone
        var listAfterResult = _tableCommands.ListTableSlicers(batch);
        Assert.DoesNotContain(listAfterResult.Slicers, s => s.Name == "SlicerToDelete");
    }

    /// <summary>
    /// Tests deleting a non-existent Table slicer returns error.
    /// </summary>
    [Fact]
    public void DeleteTableSlicer_NonExistentSlicer_ReturnsError()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Try to delete a slicer that doesn't exist
        var deleteResult = _tableCommands.DeleteTableSlicer(batch, "NonExistentSlicer");

        // Assert
        Assert.False(deleteResult.Success);
        Assert.Contains("not found", deleteResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests setting Table slicer selection for non-existent slicer returns error.
    /// </summary>
    [Fact]
    public void SetTableSlicerSelection_NonExistentSlicer_ReturnsError()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var selectionResult = _tableCommands.SetTableSlicerSelection(
            batch, "NonExistentSlicer", new List<string> { "North" });

        // Assert
        Assert.False(selectionResult.Success);
        Assert.Contains("not found", selectionResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests listing Table slicers when workbook has no slicers.
    /// </summary>
    [Fact]
    public void ListTableSlicers_NoSlicers_ReturnsEmptyList()
    {
        // Arrange - Fresh file with no slicers
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert
        Assert.True(listResult.Success, $"ListTableSlicers failed: {listResult.ErrorMessage}");
        Assert.NotNull(listResult.Slicers);
        Assert.Empty(listResult.Slicers);
    }

    /// <summary>
    /// Tests creating slicer for invalid Table column returns error.
    /// </summary>
    [Fact]
    public void CreateTableSlicer_InvalidColumn_ReturnsError()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Try to create slicer for non-existent column
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "NonExistentColumn", "InvalidSlicer", "Sales", "F2");

        // Assert
        Assert.False(slicerResult.Success);
        Assert.Contains("Column 'NonExistentColumn' not found", slicerResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests creating slicer for invalid Table name throws exception.
    /// </summary>
    [Fact]
    public void CreateTableSlicer_InvalidTableName_ReturnsError()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act & Assert - expects exception when table not found
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tableCommands.CreateTableSlicer(
                batch, "NonExistentTable", "Region", "InvalidSlicer", "Sales", "F2"));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests that Table slicer shows connected Table info.
    /// </summary>
    [Fact]
    public void CreateTableSlicer_ShowsConnectedTable()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create slicer
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "ConnectedSlicer", "Sales", "F2");

        // Assert
        Assert.True(slicerResult.Success, $"CreateTableSlicer failed: {slicerResult.ErrorMessage}");
        Assert.NotNull(slicerResult.ConnectedTable);
        Assert.Equal("SalesTable", slicerResult.ConnectedTable);
        Assert.Equal("Table", slicerResult.SourceType);
    }

    /// <summary>
    /// Tests that slicer Position is returned as a valid cell reference.
    /// This test catches bugs where Position is empty due to incorrect COM API usage.
    /// Bug context: TopLeftCell is on Slicer.Shape, not Slicer directly.
    /// </summary>
    [Fact]
    public void CreateTableSlicer_ReturnsValidPosition()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create slicer at F2
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "PositionTestSlicer", "Sales", "F2");

        // Assert - Position must be a valid cell reference, not empty
        Assert.True(slicerResult.Success, $"CreateTableSlicer failed: {slicerResult.ErrorMessage}");
        Assert.False(string.IsNullOrEmpty(slicerResult.Position),
            "Slicer Position should not be empty - verify Shape.TopLeftCell API is used correctly");
        Assert.Matches(@"^[A-Z]+\d+$", slicerResult.Position); // e.g., "F2", "AA10"
    }

    /// <summary>
    /// Tests that ListTableSlicers returns valid Position for each slicer.
    /// </summary>
    [Fact]
    public void ListTableSlicers_ReturnsValidPositionForEachSlicer()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Create slicers at different positions
        _tableCommands.CreateTableSlicer(batch, "SalesTable", "Region", "ListPosSlicer1", "Sales", "F2");
        _tableCommands.CreateTableSlicer(batch, "SalesTable", "Product", "ListPosSlicer2", "Sales", "H2");

        // Act
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert - All slicers should have valid positions
        Assert.True(listResult.Success, $"ListTableSlicers failed: {listResult.ErrorMessage}");
        foreach (var slicer in listResult.Slicers)
        {
            Assert.False(string.IsNullOrEmpty(slicer.Position),
                $"Slicer '{slicer.Name}' has empty Position - verify Shape.TopLeftCell API");
        }
    }

    /// <summary>
    /// Tests that FieldName is returned correctly (not "Unknown").
    /// This test catches bugs where SourceName property access fails silently.
    /// </summary>
    [Fact]
    public void CreateTableSlicer_ReturnsCorrectFieldName()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create slicer for "Region" column
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "FieldNameTestSlicer", "Sales", "F2");

        // Assert - FieldName must match the column name, not be "Unknown"
        Assert.True(slicerResult.Success, $"CreateTableSlicer failed: {slicerResult.ErrorMessage}");
        Assert.NotEqual("Unknown", slicerResult.FieldName);
        Assert.Equal("Region", slicerResult.FieldName);
    }

    /// <summary>
    /// Tests that ConnectedTable is returned correctly (not "Unknown" or empty).
    /// This test catches bugs where ListObject property access fails silently.
    /// </summary>
    [Fact]
    public void ListTableSlicers_ReturnsCorrectConnectedTable()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        _tableCommands.CreateTableSlicer(batch, "SalesTable", "Region", "ConnTableTestSlicer", "Sales", "F2");

        // Act
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert - ConnectedTable must be the actual table name
        Assert.True(listResult.Success, $"ListTableSlicers failed: {listResult.ErrorMessage}");
        var slicer = listResult.Slicers.FirstOrDefault(s => s.Name == "ConnTableTestSlicer");
        Assert.NotNull(slicer);
        Assert.NotEqual("Unknown", slicer.ConnectedTable);
        Assert.NotEqual(string.Empty, slicer.ConnectedTable);
        Assert.Equal("SalesTable", slicer.ConnectedTable);
    }

    /// <summary>
    /// Tests rapid sequential operations: create slicer, then immediately list slicers.
    /// This mimics MCP/LLM patterns where operations are called in rapid succession.
    /// Tests for timing issues and COM object availability.
    /// </summary>
    [Fact]
    public void RapidSequentialOperations_CreateThenList_ReturnsValidPosition()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create slicer then IMMEDIATELY list (mimics MCP agent pattern)
        var createResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "RapidTestSlicer", "Sales", "F2");
        Assert.True(createResult.Success, $"CreateTableSlicer failed: {createResult.ErrorMessage}");

        // Immediately call list - no delay (this is how MCP agents work)
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert - Both operations must succeed with valid data
        Assert.True(listResult.Success, $"ListTableSlicers failed: {listResult.ErrorMessage}");
        var slicer = listResult.Slicers.FirstOrDefault(s => s.Name == "RapidTestSlicer");
        Assert.NotNull(slicer);
        Assert.False(string.IsNullOrEmpty(slicer.Position),
            "Slicer Position empty after rapid create+list - possible COM timing issue");
        Assert.NotEqual("Unknown", slicer.FieldName);
        Assert.Equal("SalesTable", slicer.ConnectedTable);
    }

    /// <summary>
    /// Tests multiple rapid operations in sequence to stress test COM object handling.
    /// Create multiple slicers, then list all, then set selection on each.
    /// </summary>
    [Fact]
    public void RapidSequentialOperations_MultipleSlicers_AllReturnValidData()
    {
        // Arrange
        var testFile = _fixture.CreateModificationTestFile();
        using var batch = ExcelSession.BeginBatch(testFile);

        // Act - Create 3 slicers in rapid succession (using columns that exist in test data)
        var slicer1 = _tableCommands.CreateTableSlicer(batch, "SalesTable", "Region", "RapidSlicer1", "Sales", "F2");
        var slicer2 = _tableCommands.CreateTableSlicer(batch, "SalesTable", "Product", "RapidSlicer2", "Sales", "H2");
        var slicer3 = _tableCommands.CreateTableSlicer(batch, "SalesTable", "Amount", "RapidSlicer3", "Sales", "J2");

        Assert.True(slicer1.Success, $"CreateTableSlicer 1 failed: {slicer1.ErrorMessage}");
        Assert.True(slicer2.Success, $"CreateTableSlicer 2 failed: {slicer2.ErrorMessage}");
        Assert.True(slicer3.Success, $"CreateTableSlicer 3 failed: {slicer3.ErrorMessage}");

        // Immediately list all slicers
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert - All 3 slicers must have valid data
        Assert.True(listResult.Success, $"ListTableSlicers failed: {listResult.ErrorMessage}");
        Assert.True(listResult.Slicers.Count >= 3, $"Expected at least 3 slicers, got {listResult.Slicers.Count}");

        foreach (var name in new[] { "RapidSlicer1", "RapidSlicer2", "RapidSlicer3" })
        {
            var slicer = listResult.Slicers.FirstOrDefault(s => s.Name == name);
            Assert.NotNull(slicer);
            Assert.False(string.IsNullOrEmpty(slicer.Position),
                $"Slicer '{name}' has empty Position after rapid operations");
            Assert.NotEqual("Unknown", slicer.FieldName);
        }
    }

    #endregion
}




