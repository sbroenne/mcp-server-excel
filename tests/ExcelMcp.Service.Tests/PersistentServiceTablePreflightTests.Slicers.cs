using System.Globalization;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for Table Slicer operations.
/// Tests cover: create, list, set selection, delete slicers for Excel Tables.
/// Uses the per-test SalesTable baseline created by the persistent Service fixture.
/// </summary>
public sealed partial class PersistentServiceTablePreflightTests
{
    #region Table Slicer Tests

    /// <summary>
    /// Tests creating a slicer for a Table column.
    /// </summary>
    [Fact]
    public void CreateTableSlicer_ValidColumn_CreatesSlicerSuccessfully()
    {
        // Arrange - Create a fresh test file with SalesTable
        var batch = _fixture.BatchToken;
        // Act - Create slicer for Region column
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch,
            tableName: "SalesTable",
            columnName: "Region",
            slicerName: "RegionSlicer",
            destinationSheet: "Sales",
            position: "F2");

        // Assert
        RequireSuccess(slicerResult);
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
        Assert.Equal("F2", slicerResult.Position);
        Assert.Equal(["East", "North", "South", "West"], slicerResult.AvailableItems.Order());
        AssertVisibleRegions(["North", "South", "East", "West"]);
        AssertNativeSingleSlicer("RegionSlicer", "Region", "SalesTable", "Sales", "F2");
    }

    /// <summary>
    /// Tests listing Table slicers in a workbook with no filter.
    /// </summary>
    [Fact]
    public void ListTableSlicers_WithSlicers_ReturnsAllSlicers()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create two slicers
        var slicer1Result = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "RegionSlicer1", "Sales", "F2");
        RequireSuccess(slicer1Result);

        var slicer2Result = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Product", "ProductSlicer1", "Sales", "F10");
        RequireSuccess(slicer2Result);

        // Act
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert
        RequireSuccess(listResult);
        Assert.NotNull(listResult.Slicers);
        Assert.Equal(2, listResult.Slicers.Count);
        Assert.Contains(listResult.Slicers, s => s.Name == "RegionSlicer1");
        Assert.Contains(listResult.Slicers, s => s.Name == "ProductSlicer1");
        AssertVisibleRegions(["North", "South", "East", "West"]);
    }

    /// <summary>
    /// Tests listing slicers filtered by Table name.
    /// </summary>
    [Fact]
    public void ListTableSlicers_FilterByTable_ReturnsConnectedSlicersOnly()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create slicer for SalesTable
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "FilterRegionSlicer", "Sales", "F2");
        RequireSuccess(slicerResult);
        SetValues(batch, "J1:K3", [["Region", "Amount"], ["North", 7], ["South", 11]]);
        RequireSuccess(_tableCommands.Create(batch, "Sales", "OtherTable", "J1:K3"));
        RequireSuccess(_tableCommands.CreateTableSlicer(batch, "OtherTable", "Region", "OtherSlicer", "Sales", "M2"));

        // Act
        var listResult = _tableCommands.ListTableSlicers(batch, tableName: "SalesTable");

        // Assert
        RequireSuccess(listResult);
        Assert.NotNull(listResult.Slicers);
        Assert.Single(listResult.Slicers);
        Assert.Equal("FilterRegionSlicer", listResult.Slicers[0].Name);
        Assert.Equal("SalesTable", listResult.Slicers[0].ConnectedTable);
        Assert.Equal("OtherSlicer", Assert.Single(RequireSuccess(_tableCommands.ListTableSlicers(batch, "OtherTable")).Slicers).Name);
        Assert.Equal(2, RequireSuccess(_tableCommands.ListTableSlicers(batch)).Slicers.Count);
        AssertVisibleRegions(["North", "South", "East", "West"]);
    }

    /// <summary>
    /// Tests setting Table slicer selection to specific items.
    /// </summary>
    [Fact]
    public void SetTableSlicerSelection_SpecificItems_SelectsOnlyThoseItems()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create slicer
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "SelectionSlicer", "Sales", "F2");
        RequireSuccess(slicerResult);

        // Act - Select only "North" and "South"
        var selectionResult = _tableCommands.SetTableSlicerSelection(
            batch, "SelectionSlicer", new List<string> { "North", "South" }, clearFirst: true);

        // Assert
        RequireSuccess(selectionResult);
        Assert.NotNull(selectionResult.SelectedItems);
        Assert.Equal(2, selectionResult.SelectedItems.Count);
        Assert.Contains("North", selectionResult.SelectedItems);
        Assert.Contains("South", selectionResult.SelectedItems);
        Assert.DoesNotContain("East", selectionResult.SelectedItems);
        Assert.DoesNotContain("West", selectionResult.SelectedItems);
        Assert.NotNull(selectionResult.WorkflowHint);
        AssertVisibleRegions(["North", "South"]);
    }

    /// <summary>
    /// Tests clearing Table slicer selection (selecting all items).
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SetTableSlicerSelection_EmptyList_ClearsFilterSelectsAll(bool clearFirst)
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create slicer
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "ClearFilterSlicer", "Sales", "F2");
        RequireSuccess(slicerResult);

        // First, filter to just "North"
        var filtered = _tableCommands.SetTableSlicerSelection(batch, "ClearFilterSlicer", ["North"]);
        AssertTableSlicerState(filtered, 100, "North");
        AssertVisibleRegions(["North"]);

        // Act - Clear filter by passing empty list
        var clearResult = _tableCommands.SetTableSlicerSelection(
            batch, "ClearFilterSlicer", [], clearFirst);

        // Assert
        AssertTableSlicerState(clearResult, 800, "North", "South", "East", "West");
        Assert.Contains("cleared", clearResult.WorkflowHint, StringComparison.OrdinalIgnoreCase);
        AssertVisibleRegions(["North", "South", "East", "West"]);
    }

    [Fact]
    public void SetTableSlicerSelection_ClearFirstFalse_AddsWithoutReplacingAndPreservesUnfilteredState()
    {
        var batch = _fixture.BatchToken;
        var slicer = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "AddSelectionSlicer", "Sales", "F2");
        AssertTableSlicerState(slicer, 800, "North", "South", "East", "West");
        var unfiltered = _tableCommands.SetTableSlicerSelection(
            batch, "AddSelectionSlicer", ["North"], clearFirst: false);
        AssertTableSlicerState(unfiltered, 800, "North", "South", "East", "West");
        var filtered = _tableCommands.SetTableSlicerSelection(batch, "AddSelectionSlicer", ["North"]);
        AssertTableSlicerState(filtered, 100, "North");
        var added = _tableCommands.SetTableSlicerSelection(
            batch, "AddSelectionSlicer", ["South"], clearFirst: false);
        AssertTableSlicerState(added, 350, "North", "South");
        var replaced = _tableCommands.SetTableSlicerSelection(
            batch, "AddSelectionSlicer", ["South"], clearFirst: true);
        AssertTableSlicerState(replaced, 250, "South");
    }

    private void AssertTableSlicerState(SlicerResult result, double total, params string[] regions)
    {
        RequireSuccess(result);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal(["East", "North", "South", "West"], result.AvailableItems.Order());
        Assert.Equal(regions.Order(), result.SelectedItems.Order());
        var listed = _tableCommands.ListTableSlicers(_fixture.BatchToken, "SalesTable");
        RequireSuccess(listed);
        Assert.Equal(regions.Order(), Assert.Single(listed.Slicers).SelectedItems.Order());
        var visible = _tableCommands.GetData(_fixture.BatchToken, "SalesTable", visibleOnly: true);
        RequireSuccess(visible);
        Assert.Equal(regions.Length, visible.RowCount);
        Assert.Equal(regions.Length, visible.Data.Count);
        Assert.Equal(regions.Order(), visible.Data.Select(row => row[0]?.ToString()).Order());
        Assert.Equal(total, visible.Data.Sum(row => Convert.ToDouble(row[2], CultureInfo.InvariantCulture)));
    }

    [Theory]
    [InlineData("North", 100d, "South", 250d)]
    [InlineData("South", 250d, "North", 100d)]
    public void SetTableSlicerSelection_ReplaceSoleSelectedItem_SelectsOnlyReplacement(
        string initial, double initialTotal, string replacement, double replacementTotal)
    {
        var batch = _fixture.BatchToken;
        var slicer = _tableCommands.CreateTableSlicer(batch, "SalesTable", "Region",
            "ReplaceSoleItemSlicer", "Sales", "F2");
        AssertTableSlicerState(slicer, 800, "North", "South", "East", "West");
        AssertTableSlicerState(_tableCommands.SetTableSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [initial]), initialTotal, initial);
        AssertTableSlicerState(_tableCommands.SetTableSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [replacement], clearFirst: true), replacementTotal, replacement);
        AssertTableSlicerState(_tableCommands.SetTableSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [initial], clearFirst: true), initialTotal, initial);
        AssertTableSlicerState(_tableCommands.SetTableSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [replacement], clearFirst: false), 350, "North", "South");
        AssertTableSlicerState(_tableCommands.SetTableSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [replacement], clearFirst: true), replacementTotal, replacement);
        AssertTableSlicerState(_tableCommands.SetTableSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [initial.ToLowerInvariant(), initial, "missing"],
            clearFirst: true), initialTotal, initial);
        AssertTableSlicerState(_tableCommands.SetTableSlicerSelection(
            batch, "ReplaceSoleItemSlicer", []), 800, "North", "South", "East", "West");
    }

    /// <summary>
    /// Tests deleting a Table slicer from the workbook.
    /// </summary>
    [Fact]
    public void DeleteTableSlicer_ExistingSlicer_RemovesSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create slicer
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "SlicerToDelete", "Sales", "F2");
        RequireSuccess(slicerResult);
        RequireSuccess(_tableCommands.SetTableSlicerSelection(batch, "SlicerToDelete", ["South"]));
        AssertVisibleRegions(["South"]);

        // Verify slicer exists
        var listBeforeResult = RequireSuccess(_tableCommands.ListTableSlicers(batch));
        Assert.Contains(listBeforeResult.Slicers, s => s.Name == "SlicerToDelete");

        // Act
        var deleteResult = _tableCommands.DeleteTableSlicer(batch, "SlicerToDelete");

        // Assert
        RequireSuccess(deleteResult);

        // Verify slicer is gone
        var listAfterResult = RequireSuccess(_tableCommands.ListTableSlicers(batch));
        Assert.DoesNotContain(listAfterResult.Slicers, s => s.Name == "SlicerToDelete");
        Assert.Empty(listAfterResult.Slicers);
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
        AssertVisibleRegions(["South"]);
    }

    /// <summary>
    /// Tests deleting a non-existent Table slicer returns error.
    /// </summary>
    [Fact]
    public void DeleteTableSlicer_NonExistentSlicer_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var created = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "PreservedSlicer", "Sales", "F2");
        RequireSuccess(created);
        var selected = _tableCommands.SetTableSlicerSelection(batch, "PreservedSlicer", ["North"]);
        RequireSuccess(selected);
        AssertVisibleRegions(["North"]);
        var before = _tableCommands.ListTableSlicers(batch);
        RequireSuccess(before);
        Assert.Equal("PreservedSlicer", Assert.Single(before.Slicers).Name);

        // Act - Try to delete a slicer that doesn't exist
        var deleteResult = _tableCommands.DeleteTableSlicer(batch, "NonExistentSlicer");

        // Assert
        Assert.False(deleteResult.Success);
        Assert.Contains("not found", deleteResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var preserved = _tableCommands.ListTableSlicers(batch);
        RequireSuccess(preserved);
        var slicer = Assert.Single(preserved.Slicers);
        Assert.Equal("PreservedSlicer", slicer.Name);
        Assert.Equal("SalesTable", slicer.ConnectedTable);
        Assert.Equal("Region", slicer.FieldName);
        Assert.Equal(["North"], slicer.SelectedItems);
        AssertVisibleRegions(["North"]);
    }

    /// <summary>
    /// Tests setting Table slicer selection for non-existent slicer returns error.
    /// </summary>
    [Fact]
    public void SetTableSlicerSelection_NonExistentSlicer_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var created = _tableCommands.CreateTableSlicer(batch, "SalesTable", "Region", "RetainedSelectionSlicer", "Sales", "F2");
        RequireSuccess(created);
        var selected = _tableCommands.SetTableSlicerSelection(batch, "RetainedSelectionSlicer", ["South"]);
        RequireSuccess(selected);
        AssertVisibleRegions(["South"]);

        // Act
        var selectionResult = _tableCommands.SetTableSlicerSelection(
            batch, "NonExistentSlicer", new List<string> { "North" });

        // Assert
        Assert.False(selectionResult.Success);
        Assert.Contains("not found", selectionResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        AssertVisibleRegions(["South"]);
        var listed = _tableCommands.ListTableSlicers(batch);
        RequireSuccess(listed);
        Assert.Equal(["South"], Assert.Single(listed.Slicers).SelectedItems);
    }

    [Fact]
    public void SetTableSlicerSelection_AddToSelection_PreservesPreviouslyVisibleRows()
    {
        var batch = _fixture.BatchToken;
        var created = _tableCommands.CreateTableSlicer(batch, "SalesTable", "Region", "AdditiveSlicer", "Sales", "F2");
        RequireSuccess(created);
        var first = _tableCommands.SetTableSlicerSelection(batch, "AdditiveSlicer", ["North"]);
        RequireSuccess(first);
        AssertVisibleRegions(["North"]);

        var added = _tableCommands.SetTableSlicerSelection(batch, "AdditiveSlicer", ["South"], clearFirst: false);

        RequireSuccess(added);
        AssertVisibleRegions(["North", "South"]);
        var listed = _tableCommands.ListTableSlicers(batch);
        RequireSuccess(listed);
        Assert.Equal(["North", "South"], Assert.Single(listed.Slicers).SelectedItems.Order());
    }

    /// <summary>
    /// Tests listing Table slicers when workbook has no slicers.
    /// </summary>
    [Fact]
    public void ListTableSlicers_NoSlicers_ReturnsEmptyList()
    {
        // Arrange - Fresh file with no slicers
        var batch = _fixture.BatchToken;

        // Act
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert
        RequireSuccess(listResult);
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
        var batch = _fixture.BatchToken;

        // Act - Try to create slicer for non-existent column
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "NonExistentColumn", "InvalidSlicer", "Sales", "F2");

        // Assert
        Assert.False(slicerResult.Success);
        Assert.Contains("Column 'NonExistentColumn' not found", slicerResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        var listed = _tableCommands.ListTableSlicers(batch);
        RequireSuccess(listed);
        Assert.Empty(listed.Slicers);
        AssertVisibleRegions(["North", "South", "East", "West"]);
    }

    /// <summary>
    /// Tests creating slicer for invalid Table name throws exception.
    /// </summary>
    [Fact]
    public void CreateTableSlicer_InvalidTableName_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act & Assert - expects exception when table not found
        RequireSuccess(_tableCommands.CreateTableSlicer(batch, "SalesTable", "Region", "KeptSlicer", "Sales", "F2"));
        RequireSuccess(_tableCommands.SetTableSlicerSelection(batch, "KeptSlicer", ["East"]));
        AssertVisibleRegions(["East"]);
        var ex = Assert.Throws<InvalidOperationException>(() =>
            _tableCommands.CreateTableSlicer(
                batch, "NonExistentTable", "Region", "InvalidSlicer", "Sales", "F2"));
        Assert.Contains("not found", ex.Message, StringComparison.OrdinalIgnoreCase);
        var retained = Assert.Single(RequireSuccess(_tableCommands.ListTableSlicers(batch)).Slicers);
        Assert.Equal("KeptSlicer", retained.Name);
        Assert.Equal(["East"], retained.SelectedItems);
        AssertVisibleRegions(["East"]);
    }

    /// <summary>
    /// Tests that Table slicer shows connected Table info.
    /// </summary>
    [Fact]
    public void CreateTableSlicer_ShowsConnectedTable()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act - Create slicer
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "ConnectedSlicer", "Sales", "F2");

        // Assert
        RequireSuccess(slicerResult);
        Assert.NotNull(slicerResult.ConnectedTable);
        Assert.Equal("SalesTable", slicerResult.ConnectedTable);
        Assert.Equal("Table", slicerResult.SourceType);
        Assert.Equal("ConnectedSlicer", Assert.Single(RequireSuccess(_tableCommands.ListTableSlicers(batch)).Slicers).Name);
        AssertVisibleRegions(["North", "South", "East", "West"]);
        AssertNativeSingleSlicer("ConnectedSlicer", "Region", "SalesTable", "Sales", "F2");
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
        var batch = _fixture.BatchToken;

        // Act - Create slicer at F2
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "PositionTestSlicer", "Sales", "F2");

        // Assert - Position must be a valid cell reference, not empty
        RequireSuccess(slicerResult);
        Assert.False(string.IsNullOrEmpty(slicerResult.Position),
            "Slicer Position should not be empty - verify Shape.TopLeftCell API is used correctly");
        Assert.Matches(@"^[A-Z]+\d+$", slicerResult.Position); // e.g., "F2", "AA10"
        Assert.Equal("F2", slicerResult.Position);
        AssertNativeSingleSlicer("PositionTestSlicer", "Region", "SalesTable", "Sales", "F2");
    }

    /// <summary>
    /// Tests that ListTableSlicers returns valid Position for each slicer.
    /// </summary>
    [Fact]
    public void ListTableSlicers_ReturnsValidPositionForEachSlicer()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create slicers at different positions
        RequireSuccess(_tableCommands.CreateTableSlicer(batch, "SalesTable", "Region", "ListPosSlicer1", "Sales", "F2"));
        RequireSuccess(_tableCommands.CreateTableSlicer(batch, "SalesTable", "Product", "ListPosSlicer2", "Sales", "H2"));

        // Act
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert - All slicers should have valid positions
        RequireSuccess(listResult);
        Assert.Equal(2, listResult.Slicers.Count);
        Assert.Equal("F2", Assert.Single(listResult.Slicers, slicer => slicer.Name == "ListPosSlicer1").Position);
        Assert.Equal("H2", Assert.Single(listResult.Slicers, slicer => slicer.Name == "ListPosSlicer2").Position);
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
        var batch = _fixture.BatchToken;

        // Act - Create slicer for "Region" column
        var slicerResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "FieldNameTestSlicer", "Sales", "F2");

        // Assert - FieldName must match the column name, not be "Unknown"
        RequireSuccess(slicerResult);
        Assert.NotEqual("Unknown", slicerResult.FieldName);
        Assert.Equal("Region", slicerResult.FieldName);
        AssertNativeSingleSlicer("FieldNameTestSlicer", "Region", "SalesTable", "Sales", "F2");
    }

    /// <summary>
    /// Tests that ConnectedTable is returned correctly (not "Unknown" or empty).
    /// This test catches bugs where ListObject property access fails silently.
    /// </summary>
    [Fact]
    public void ListTableSlicers_ReturnsCorrectConnectedTable()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        RequireSuccess(_tableCommands.CreateTableSlicer(batch, "SalesTable", "Region", "ConnTableTestSlicer", "Sales", "F2"));

        // Act
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert - ConnectedTable must be the actual table name
        RequireSuccess(listResult);
        var slicer = Assert.Single(listResult.Slicers);
        Assert.Equal("ConnTableTestSlicer", slicer.Name);
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
        var batch = _fixture.BatchToken;

        // Act - Create slicer then IMMEDIATELY list (mimics MCP agent pattern)
        var createResult = _tableCommands.CreateTableSlicer(
            batch, "SalesTable", "Region", "RapidTestSlicer", "Sales", "F2");
        RequireSuccess(createResult);

        // Immediately call list - no delay (this is how MCP agents work)
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert - Both operations must succeed with valid data
        RequireSuccess(listResult);
        var slicer = Assert.Single(listResult.Slicers);
        Assert.Equal("RapidTestSlicer", slicer.Name);
        Assert.False(string.IsNullOrEmpty(slicer.Position),
            "Slicer Position empty after rapid create+list - possible COM timing issue");
        Assert.Equal("F2", slicer.Position);
        Assert.Equal("Region", slicer.FieldName);
        Assert.Equal("SalesTable", slicer.ConnectedTable);
        AssertVisibleRegions(["North", "South", "East", "West"]);
    }

    /// <summary>
    /// Tests multiple rapid operations in sequence to stress test COM object handling.
    /// Create multiple slicers, then list all, then set selection on each.
    /// </summary>
    [Fact]
    public void RapidSequentialOperations_MultipleSlicers_AllReturnValidData()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Act - Create 3 slicers in rapid succession (using columns that exist in test data)
        var slicer1 = _tableCommands.CreateTableSlicer(batch, "SalesTable", "Region", "RapidSlicer1", "Sales", "F2");
        var slicer2 = _tableCommands.CreateTableSlicer(batch, "SalesTable", "Product", "RapidSlicer2", "Sales", "H2");
        var slicer3 = _tableCommands.CreateTableSlicer(batch, "SalesTable", "Amount", "RapidSlicer3", "Sales", "J2");

        RequireSuccess(slicer1);
        RequireSuccess(slicer2);
        RequireSuccess(slicer3);

        // Immediately list all slicers
        var listResult = _tableCommands.ListTableSlicers(batch);

        // Assert - All 3 slicers must have valid data
        RequireSuccess(listResult);
        Assert.Equal(3, listResult.Slicers.Count);

        foreach (var (name, field, position) in new[]
            { ("RapidSlicer1", "Region", "F2"), ("RapidSlicer2", "Product", "H2"), ("RapidSlicer3", "Amount", "J2") })
        {
            var slicer = Assert.Single(listResult.Slicers, item => item.Name == name);
            Assert.Equal(position, slicer.Position);
            Assert.Equal(field, slicer.FieldName);
            Assert.Equal("SalesTable", slicer.ConnectedTable);
        }
        AssertVisibleRegions(["North", "South", "East", "West"]);
    }

    #endregion

    private void AssertNativeSingleSlicer(string name, string field, string tableName, string sheetName, string position)
    {
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.SlicerCaches? caches = null;
            Excel.SlicerCache? cache = null;
            Excel.Slicers? slicers = null;
            Excel.Slicer? slicer = null;
            Excel.ListObject? table = null;
            Excel.Shape? shape = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? anchor = null;
            try
            {
                caches = context.Book.SlicerCaches;
                Assert.Equal(1, caches.Count);
                cache = caches.Item[1];
                Assert.True(cache.List);
                Assert.Equal(field, cache.SourceName);
                table = cache.ListObject;
                Assert.Equal(tableName, table.Name);
                slicers = cache.Slicers;
                Assert.Equal(1, slicers.Count);
                slicer = slicers.Item[1];
                Assert.Equal(name, slicer.Name);
                shape = slicer.Shape;
                sheet = (Excel.Worksheet)shape.Parent;
                Assert.Equal(sheetName, sheet.Name);
                anchor = shape.TopLeftCell;
                Assert.Equal(position, anchor.Address[false, false]);
            }
            finally
            {
                ComUtilities.Release(ref anchor);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref shape);
                ComUtilities.Release(ref slicer);
                ComUtilities.Release(ref slicers);
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref cache);
                ComUtilities.Release(ref caches);
            }
        });
    }

    private void AssertVisibleRegions(string[] expected)
    {
        var visible = RequireSuccess(_tableCommands.GetData(_fixture.BatchToken, "SalesTable", visibleOnly: true));
        RequireSuccess(visible);
        Assert.Equal(expected, visible.Data.Select(row => row[0]?.ToString()));
        var all = RequireSuccess(_tableCommands.GetData(_fixture.BatchToken, "SalesTable"));
        RequireSuccess(all);
        Assert.Equal(["North", "South", "East", "West"], all.Data.Select(row => row[0]?.ToString()));
        Assert.Equal([100, 250, 150, 300], all.Data.Select(row =>
            Convert.ToInt32(row[2], System.Globalization.CultureInfo.InvariantCulture)));
        AssertSalesData(all);
        Assert.Equal(all.Headers, visible.Headers);
        AssertSalesRows(visible.Data, expected.Select(region =>
            all.Data.FindIndex(row => Equals(row[0], region))).ToArray());
    }
}
