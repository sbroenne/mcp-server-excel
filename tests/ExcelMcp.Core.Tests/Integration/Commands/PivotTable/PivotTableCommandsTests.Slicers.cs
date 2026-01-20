using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PivotTable;

public partial class PivotTableCommandsTests
{
    /// <summary>
    /// Tests creating a slicer for a PivotTable field.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void CreateSlicer_ValidField_CreatesSlicerSuccessfully()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(CreateSlicer_ValidField_CreatesSlicerSuccessfully));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "SlicerTest");
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
            destinationSheet: "SalesData",
            position: "I2");

        // Assert
        Assert.True(slicerResult.Success, $"CreateSlicer failed: {slicerResult.ErrorMessage}");
        Assert.Equal("RegionSlicer", slicerResult.Name);
        Assert.Equal("Region", slicerResult.FieldName);
        Assert.Equal("SalesData", slicerResult.SheetName);
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
        var testFile = CreateTestFileWithData(nameof(ListSlicers_WithSlicers_ReturnsAllSlicers));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "ListSlicersTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add fields
        _pivotCommands.AddRowField(batch, "ListSlicersTest", "Region");
        _pivotCommands.AddRowField(batch, "ListSlicersTest", "Product");

        // Create two slicers
        var slicer1Result = _pivotCommands.CreateSlicer(
            batch, "ListSlicersTest", "Region", "RegionSlicer1", "SalesData", "I2");
        Assert.True(slicer1Result.Success, $"Failed to create slicer 1: {slicer1Result.ErrorMessage}");

        var slicer2Result = _pivotCommands.CreateSlicer(
            batch, "ListSlicersTest", "Product", "ProductSlicer1", "SalesData", "I10");
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
        var testFile = CreateTestFileWithData(nameof(ListSlicers_FilterByPivotTable_ReturnsConnectedSlicersOnly));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "FilterSlicersTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        // Add Region field and create slicer
        _pivotCommands.AddRowField(batch, "FilterSlicersTest", "Region");
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "FilterSlicersTest", "Region", "FilterRegionSlicer", "SalesData", "I2");
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
        var testFile = CreateTestFileWithData(nameof(SetSlicerSelection_SpecificItems_SelectsOnlyThoseItems));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable with Region field
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "SelectionTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "SelectionTest", "Region");

        // Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "SelectionTest", "Region", "SelectionSlicer", "SalesData", "I2");
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
    [Fact]
    [Trait("Speed", "Medium")]
    public void SetSlicerSelection_EmptyList_ClearsFilterSelectsAll()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(SetSlicerSelection_EmptyList_ClearsFilterSelectsAll));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "ClearFilterTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "ClearFilterTest", "Region");

        // Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "ClearFilterTest", "Region", "ClearFilterSlicer", "SalesData", "I2");
        Assert.True(slicerResult.Success, $"Failed to create slicer: {slicerResult.ErrorMessage}");

        // First, filter to just "North"
        _pivotCommands.SetSlicerSelection(batch, "ClearFilterSlicer", new List<string> { "North" });

        // Act - Clear filter by passing empty list
        var clearResult = _pivotCommands.SetSlicerSelection(
            batch, "ClearFilterSlicer", new List<string>());

        // Assert
        Assert.True(clearResult.Success, $"SetSlicerSelection (clear) failed: {clearResult.ErrorMessage}");
        Assert.NotNull(clearResult.SelectedItems);
        Assert.True(clearResult.SelectedItems.Count >= 2, "Expected all items to be selected after clear");
        Assert.Contains("North", clearResult.SelectedItems);
        Assert.Contains("South", clearResult.SelectedItems);
        Assert.Contains("cleared", clearResult.WorkflowHint, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Tests deleting a slicer from the workbook.
    /// </summary>
    [Fact]
    [Trait("Speed", "Medium")]
    public void DeleteSlicer_ExistingSlicer_RemovesSuccessfully()
    {
        // Arrange
        var testFile = CreateTestFileWithData(nameof(DeleteSlicer_ExistingSlicer_RemovesSuccessfully));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "DeleteSlicerTest");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "DeleteSlicerTest", "Region");

        // Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "DeleteSlicerTest", "Region", "SlicerToDelete", "SalesData", "I2");
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
        var testFile = CreateTestFileWithData(nameof(DeleteSlicer_NonExistentSlicer_ReturnsError));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable (no slicer)
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "NoSlicerTest");
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
        var testFile = CreateTestFileWithData(nameof(SetSlicerSelection_NonExistentSlicer_ReturnsError));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable (no slicer)
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "NoSlicerTest2");
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
        var testFile = CreateTestFileWithData(nameof(ListSlicers_NoSlicers_ReturnsEmptyList));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable without any slicers
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "NoSlicerPivot");
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
        var testFile = CreateTestFileWithData(nameof(CreateSlicer_ShowsConnectedPivotTable));

        var logger = _loggerFactory.CreateLogger<ExcelBatch>();
        using var batch = new ExcelBatch(new[] { testFile }, logger);

        // Create PivotTable
        var createResult = _pivotCommands.CreateFromRange(
            batch, "SalesData", "A1:D6", "SalesData", "F2", "ConnectedPivot");
        Assert.True(createResult.Success, $"Failed to create PivotTable: {createResult.ErrorMessage}");

        _pivotCommands.AddRowField(batch, "ConnectedPivot", "Region");

        // Act - Create slicer
        var slicerResult = _pivotCommands.CreateSlicer(
            batch, "ConnectedPivot", "Region", "ConnectedSlicer", "SalesData", "I2");

        // Assert
        Assert.True(slicerResult.Success, $"CreateSlicer failed: {slicerResult.ErrorMessage}");
        Assert.NotNull(slicerResult.ConnectedPivotTables);
        Assert.Contains("ConnectedPivot", slicerResult.ConnectedPivotTables);
    }
}
