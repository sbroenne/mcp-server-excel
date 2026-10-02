using System.Globalization;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServicePivotTableTests
{
    /// <summary>
    /// Tests creating a slicer for a PivotTable field.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateSlicer_ValidField_CreatesSlicerSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SlicerTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add Region to Row area
        var addFieldResult = _pivotCommands.AddRowField(batch, "SlicerTest", "Region");
        Assert.True(addFieldResult.Success, $"Failed to add Region field: {addFieldResult.ErrorMessage}");

        // Act - Create slicer for Region field
        var slicerResult = _pivotCommands.CreateSlicer(
            batch,
            pivotTableName: "SlicerTest",
            fieldName: "Region",
            slicerName: "RegionSlicer",
            destinationSheet: _salesSheetName,
            position: "I2");

        // Assert
        Assert.True(slicerResult.Success, $"CreateSlicer failed: {slicerResult.ErrorMessage}");
        Assert.Equal("RegionSlicer", slicerResult.Name);
        Assert.Equal("Region", slicerResult.FieldName);
        Assert.Equal(_salesSheetName, slicerResult.SheetName);
        Assert.NotNull(slicerResult.AvailableItems);
        Assert.Contains("North", slicerResult.AvailableItems);
        Assert.Contains("South", slicerResult.AvailableItems);
        Assert.NotNull(slicerResult.WorkflowHint);
    }

    /// <summary>
    /// Tests listing slicers in a workbook with no filter.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void ListSlicers_WithSlicers_ReturnsAllSlicers()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ListSlicersTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add fields
        _pivotCommands.AddRowField(batch, "ListSlicersTest", "Region");
        _pivotCommands.AddRowField(batch, "ListSlicersTest", "Product");

        // Create two slicers
        var slicer1Result = _pivotCommands.CreateSlicer(
            batch, "ListSlicersTest", "Region", "RegionSlicer1", _salesSheetName, "I2");
        Assert.True(slicer1Result.Success, $"Failed to create slicer 1: {slicer1Result.ErrorMessage}");

        var slicer2Result = _pivotCommands.CreateSlicer(
            batch, "ListSlicersTest", "Product", "ProductSlicer1", _salesSheetName, "I10");
        Assert.True(slicer2Result.Success, $"Failed to create slicer 2: {slicer2Result.ErrorMessage}");

        // Act
        var listResult = _pivotCommands.ListSlicers(batch);

        // Assert
        Assert.True(listResult.Success, $"ListSlicers failed: {listResult.ErrorMessage}");
        Assert.NotNull(listResult.Slicers);
        Assert.True(listResult.Slicers.Count >= 2, $"Expected at least 2 slicers, got {listResult.Slicers.Count}");
        Assert.Contains(listResult.Slicers, s => s.Name == "RegionSlicer1");
        Assert.Contains(listResult.Slicers, s => s.Name == "ProductSlicer1");
    }

    /// <summary>
    /// Tests listing slicers filtered by PivotTable name.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void ListSlicers_FilterByPivotTable_ReturnsConnectedSlicersOnly()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "FilterSlicersTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add Region field and create slicer
        _pivotCommands.AddRowField(batch, "FilterSlicersTest", "Region");
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "FilterSlicersTest", "Region", "FilterRegionSlicer", _salesSheetName, "I2");
        Assert.True(slicerResult.Success, $"Failed to create slicer: {slicerResult.ErrorMessage}");

        // Act
        var listResult = _pivotCommands.ListSlicers(batch, pivotTableName: "FilterSlicersTest");

        // Assert
        Assert.True(listResult.Success, $"ListSlicers failed: {listResult.ErrorMessage}");
        Assert.NotNull(listResult.Slicers);
        Assert.Single(listResult.Slicers);
        Assert.Equal("FilterRegionSlicer", listResult.Slicers[0].Name);
    }

    /// <summary>
    /// Tests setting slicer selection to specific items.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_SpecificItems_SelectsOnlyThoseItems()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable with Region field
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "SelectionTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "SelectionTest", "Region");

        // Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "SelectionTest", "Region", "SelectionSlicer", _salesSheetName, "I2");
        Assert.True(slicerResult.Success, $"Failed to create slicer: {slicerResult.ErrorMessage}");

        // Act - Select only "North"
        var selectionResult = _pivotCommands.SetSlicerSelection(
            batch, "SelectionSlicer", new List<string> { "North" }, clearFirst: true);

        // Assert
        Assert.True(selectionResult.Success, $"SetSlicerSelection failed: {selectionResult.ErrorMessage}");
        Assert.NotNull(selectionResult.SelectedItems);
        Assert.Single(selectionResult.SelectedItems);
        Assert.Contains("North", selectionResult.SelectedItems);
        Assert.DoesNotContain("South", selectionResult.SelectedItems);
        Assert.NotNull(selectionResult.WorkflowHint);
    }

    /// <summary>
    /// Tests clearing slicer selection (selecting all items).
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_EmptyList_ClearsFilterSelectsAll(bool clearFirst)
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ClearFilterTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        Assert.True(_pivotCommands.AddRowField(batch, "ClearFilterTest", "Region").Success);
        Assert.True(_pivotCommands.AddValueField(batch, "ClearFilterTest", "Sales").Success);

        // Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "ClearFilterTest", "Region", "ClearFilterSlicer", _salesSheetName, "I2");
        Assert.True(slicerResult.Success, $"Failed to create slicer: {slicerResult.ErrorMessage}");

        // First, filter to just "North"
        var filtered = _pivotCommands.SetSlicerSelection(batch, "ClearFilterSlicer", ["North"]);
        AssertRegularSlicerState(filtered, 325, "North");

        // Act - Clear filter by passing empty list
        var clearResult = _pivotCommands.SetSlicerSelection(
            batch, "ClearFilterSlicer", [], clearFirst);

        // Assert
        AssertRegularSlicerState(clearResult, 650, "North", "South");
        Assert.Contains("cleared", clearResult.WorkflowHint, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_ClearFirstFalse_AddsWithoutReplacingAndPreservesUnfilteredState()
    {
        var batch = _fixture.BatchToken;
        var created = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "AddSelectionTest");
        Assert.True(created.Success, created.ErrorMessage);
        Assert.True(_pivotCommands.AddRowField(batch, "AddSelectionTest", "Region").Success);
        Assert.True(_pivotCommands.AddValueField(batch, "AddSelectionTest", "Sales").Success);
        var slicer = _pivotCommands.CreateSlicer(
            batch, "AddSelectionTest", "Region", "AddSelectionSlicer", _salesSheetName, "I2");
        AssertRegularSlicerState(slicer, 650, "North", "South");
        var unfiltered = _pivotCommands.SetSlicerSelection(
            batch, "AddSelectionSlicer", ["North"], clearFirst: false);
        AssertRegularSlicerState(unfiltered, 650, "North", "South");
        var filtered = _pivotCommands.SetSlicerSelection(batch, "AddSelectionSlicer", ["North"]);
        AssertRegularSlicerState(filtered, 325, "North");
        var added = _pivotCommands.SetSlicerSelection(
            batch, "AddSelectionSlicer", ["South"], clearFirst: false);
        AssertRegularSlicerState(added, 650, "North", "South");
        var replaced = _pivotCommands.SetSlicerSelection(
            batch, "AddSelectionSlicer", ["South"], clearFirst: true);
        AssertRegularSlicerState(replaced, 325, "South");
    }

    private void AssertRegularSlicerState(SlicerResult result, double total, params string[] regions)
    {
        Assert.True(result.Success, result.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal(["North", "South"], result.AvailableItems.Order());
        Assert.Equal(regions.Order(), result.SelectedItems.Order());
        var listed = _pivotCommands.ListSlicers(_fixture.BatchToken);
        Assert.True(listed.Success, listed.ErrorMessage);
        Assert.Equal(regions.Order(), Assert.Single(listed.Slicers).SelectedItems.Order());
        var values = _commands.GetValues(_fixture.BatchToken, _salesSheetName, "F2:G6");
        Assert.True(values.Success, values.ErrorMessage);
        var rows = values.Values.Where(row => row[0] is not null
            && double.TryParse(row[1]?.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out _)).ToList();
        Assert.Equal(regions.Length + 1, rows.Count);
        Assert.Equal(regions.Order(), rows.Take(regions.Length).Select(row => row[0]!.ToString()).Order());
        Assert.All(rows.Take(regions.Length), row =>
            Assert.Equal(325, double.Parse(row[1]!.ToString()!, CultureInfo.InvariantCulture)));
        Assert.Equal(total, double.Parse(rows[^1][1]!.ToString()!, CultureInfo.InvariantCulture));
    }

    [Theory]
    [InlineData("North", "South")]
    [InlineData("South", "North")]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_ReplaceSoleSelectedItem_SelectsOnlyReplacement(
        string initial, string replacement)
    {
        var batch = _fixture.BatchToken;
        var created = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ReplaceSoleItemPivot");
        Assert.True(created.Success, created.ErrorMessage);
        Assert.True(_pivotCommands.AddRowField(batch, "ReplaceSoleItemPivot", "Region").Success);
        Assert.True(_pivotCommands.AddValueField(batch, "ReplaceSoleItemPivot", "Sales").Success);
        var slicer = _pivotCommands.CreateSlicer(batch, "ReplaceSoleItemPivot", "Region",
            "ReplaceSoleItemSlicer", _salesSheetName, "I2");
        AssertRegularSlicerState(slicer, 650, "North", "South");

        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [initial]), 325, initial);
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [replacement], clearFirst: true), 325, replacement);
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [initial], clearFirst: true), 325, initial);
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [replacement], clearFirst: false), 650, "North", "South");
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [replacement], clearFirst: true), 325, replacement);
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", [initial.ToLowerInvariant(), initial, "missing"],
            clearFirst: true), 325, initial);
        AssertRegularSlicerState(_pivotCommands.SetSlicerSelection(
            batch, "ReplaceSoleItemSlicer", []), 650, "North", "South");
    }

    /// <summary>
    /// Tests deleting a slicer from the workbook.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteSlicer_ExistingSlicer_RemovesSuccessfully()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "DeleteSlicerTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "DeleteSlicerTest", "Region");

        // Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "DeleteSlicerTest", "Region", "SlicerToDelete", _salesSheetName, "I2");
        Assert.True(slicerResult.Success, $"Failed to create slicer: {slicerResult.ErrorMessage}");

        // Verify slicer exists
        var listBeforeResult = _pivotCommands.ListSlicers(batch);
        Assert.Contains(listBeforeResult.Slicers, s => s.Name == "SlicerToDelete");

        // Act
        var deleteResult = _pivotCommands.DeleteSlicer(batch, "SlicerToDelete");

        // Assert
        Assert.True(deleteResult.Success, $"DeleteSlicer failed: {deleteResult.ErrorMessage}");

        // Verify slicer is gone
        var listAfterResult = _pivotCommands.ListSlicers(batch);
        Assert.DoesNotContain(listAfterResult.Slicers, s => s.Name == "SlicerToDelete");
    }

    /// <summary>
    /// Tests deleting a non-existent slicer returns error.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteSlicer_NonExistentSlicer_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable (no slicer)
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "NoSlicerTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Act
        var deleteResult = _pivotCommands.DeleteSlicer(batch, "NonExistentSlicer");

        // Assert
        Assert.False(deleteResult.Success);
        Assert.Contains("not found", deleteResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests setting slicer selection for non-existent slicer returns error.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_NonExistentSlicer_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable (no slicer)
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "NoSlicerTest2");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Act
        var selectionResult = _pivotCommands.SetSlicerSelection(
            batch, "NonExistentSlicer", new List<string> { "North" });

        // Assert
        Assert.False(selectionResult.Success);
        Assert.Contains("not found", selectionResult.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests listing slicers when workbook has no slicers.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void ListSlicers_NoSlicers_ReturnsEmptyList()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable without any slicers
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "NoSlicerPivot");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Act
        var listResult = _pivotCommands.ListSlicers(batch);

        // Assert
        Assert.True(listResult.Success, $"ListSlicers failed: {listResult.ErrorMessage}");
        Assert.NotNull(listResult.Slicers);
        Assert.Empty(listResult.Slicers);
    }

    /// <summary>
    /// Tests that slicer shows connected PivotTables.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateSlicer_ShowsConnectedPivotTable()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ConnectedPivot");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "ConnectedPivot", "Region");

        // Act - Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "ConnectedPivot", "Region", "ConnectedSlicer", _salesSheetName, "I2");

        // Assert
        Assert.True(slicerResult.Success, $"CreateSlicer failed: {slicerResult.ErrorMessage}");
        Assert.NotNull(slicerResult.ConnectedPivotTables);
        Assert.Contains("ConnectedPivot", slicerResult.ConnectedPivotTables);
    }

    /// <summary>
    /// Tests that slicer Position is returned as a valid cell reference.
    /// This test catches bugs where Position is empty due to incorrect COM API usage.
    /// Bug context: TopLeftCell is on Slicer.Shape, not Slicer directly.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateSlicer_ReturnsValidPosition()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "PositionTestPivot");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");
        _pivotCommands.AddRowField(batch, "PositionTestPivot", "Region");

        // Act - Create slicer at I2
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "PositionTestPivot", "Region", "PositionTestSlicer", _salesSheetName, "I2");

        // Assert - Position must be a valid cell reference, not empty
        Assert.True(slicerResult.Success, $"CreateSlicer failed: {slicerResult.ErrorMessage}");
        Assert.False(string.IsNullOrEmpty(slicerResult.Position),
            "Slicer Position should not be empty - verify Shape.TopLeftCell API is used correctly");
        Assert.Matches(@"^[A-Z]+\d+$", slicerResult.Position); // e.g., "I2", "AA10"
    }

    /// <summary>
    /// Tests that ListSlicers returns valid Position for each slicer.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void ListSlicers_ReturnsValidPositionForEachSlicer()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable with two slicers
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ListPosTestPivot");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");
        _pivotCommands.AddRowField(batch, "ListPosTestPivot", "Region");
        _pivotCommands.AddRowField(batch, "ListPosTestPivot", "Product");

        _pivotCommands.CreateSlicer(batch, "ListPosTestPivot", "Region", "ListPosSlicer1", _salesSheetName, "I2");
        _pivotCommands.CreateSlicer(batch, "ListPosTestPivot", "Product", "ListPosSlicer2", _salesSheetName, "K2");

        // Act
        var listResult = _pivotCommands.ListSlicers(batch);

        // Assert - All slicers should have valid positions
        Assert.True(listResult.Success, $"ListSlicers failed: {listResult.ErrorMessage}");
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
    [Trait("Speed", "Medium")]
    public void CreateSlicer_ReturnsCorrectFieldName()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "FieldNameTestPivot");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");
        _pivotCommands.AddRowField(batch, "FieldNameTestPivot", "Region");

        // Act - Create slicer for "Region" field
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "FieldNameTestPivot", "Region", "FieldNameTestSlicer", _salesSheetName, "I2");

        // Assert - FieldName must match the field name, not be "Unknown"
        Assert.True(slicerResult.Success, $"CreateSlicer failed: {slicerResult.ErrorMessage}");
        Assert.NotEqual("Unknown", slicerResult.FieldName);
        Assert.Equal("Region", slicerResult.FieldName);
    }

    /// <summary>
    /// Tests that ConnectedPivotTables is returned correctly (not empty).
    /// This test catches bugs where PivotTables collection access fails silently.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void ListSlicers_ReturnsCorrectConnectedPivotTables()
    {
        // Arrange
        var batch = _fixture.BatchToken;

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, _salesSheetName, "A1:D6", _salesSheetName, "F2", "ConnPivotTestPivot");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");
        _pivotCommands.AddRowField(batch, "ConnPivotTestPivot", "Region");
        _pivotCommands.CreateSlicer(batch, "ConnPivotTestPivot", "Region", "ConnPivotTestSlicer", _salesSheetName, "I2");

        // Act
        var listResult = _pivotCommands.ListSlicers(batch);

        // Assert - ConnectedPivotTables must contain our PivotTable
        Assert.True(listResult.Success, $"ListSlicers failed: {listResult.ErrorMessage}");
        var slicer = listResult.Slicers.FirstOrDefault(s => s.Name == "ConnPivotTestSlicer");
        Assert.NotNull(slicer);
        Assert.NotNull(slicer.ConnectedPivotTables);
        Assert.NotEmpty(slicer.ConnectedPivotTables);
        Assert.Contains("ConnPivotTestPivot", slicer.ConnectedPivotTables);
    }
}

